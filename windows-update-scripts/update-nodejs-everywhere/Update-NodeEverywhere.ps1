<#
.SYNOPSIS
Updates Node.js everywhere on a Windows server:
- NVM-managed Node versions (per installed major branch)
- Non-NVM system Node install (winget first, MSI fallback)

.BEHAVIOR
- Logs all actions to console and log file
- Continues executing remaining steps even if one fails
- Shows final summary + pauses for review
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$UpdateNvm = $true,
    [switch]$CleanupOldNvmVersions,
    [switch]$UpdateSystemInstall = $true,
    [ValidateSet("LTS", "Current")] [string]$Channel = "LTS",
    [switch]$UseWinget = $true,
    [string]$LogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $LogDirectory) {
    $LogDirectory = Join-Path -Path $PSScriptRoot -ChildPath "logs"
}

# -----------------------------
# Logging + step tracking
# -----------------------------
$script:StepResults = New-Object System.Collections.Generic.List[object]

if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$script:LogFile = Join-Path $LogDirectory ("node-update-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")] [string]$Level = "INFO"
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] [{1}] {2}" -f $ts, $Level, $Message
    Add-Content -Path $script:LogFile -Value $line

    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

function Invoke-Step {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [scriptblock]$Action
    )

    Write-Log "START: $Name"
    try {
        & $Action
        $script:StepResults.Add([pscustomobject]@{ Step = $Name; Status = "SUCCESS"; Error = "" })
        Write-Log "END: $Name" "SUCCESS"
    }
    catch {
        $msg = $_.Exception.Message
        $script:StepResults.Add([pscustomobject]@{ Step = $Name; Status = "FAILED"; Error = $msg })
        Write-Log "FAILED: $Name :: $msg" "ERROR"
    }
}

function Invoke-CmdText {
    param([Parameter(Mandatory)] [string]$Command)

    Write-Log "CMD: $Command"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $p.StartInfo.FileName = "cmd.exe"
    $p.StartInfo.Arguments = "/d /c $Command"
    $p.StartInfo.RedirectStandardOutput = $true
    $p.StartInfo.RedirectStandardError = $true
    $p.StartInfo.UseShellExecute = $false
    $p.StartInfo.CreateNoWindow = $true

    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    if ($stdout) { Add-Content -Path $script:LogFile -Value $stdout.TrimEnd() }
    if ($stderr) { Add-Content -Path $script:LogFile -Value $stderr.TrimEnd() }

    if ($p.ExitCode -ne 0) {
        throw "Command failed (exit $($p.ExitCode)): $Command"
    }

    return ($stdout -split "`r?`n")
}

# -----------------------------
# Node/NVM helpers
# -----------------------------
function Get-NvmVersions {
    $lines = Invoke-CmdText -Command "nvm ls"
    $versions = @()

    foreach ($line in $lines) {
        if ($line -match '^\s*(\*)?\s*([0-9]+\.[0-9]+\.[0-9]+)') {
            $isCurrent = $matches[1] -eq '*'
            $verText = $matches[2]
            $ver = [version]$verText

            $versions += [pscustomobject]@{
                VersionText = $verText
                Version     = $ver
                Major       = $ver.Major
                IsCurrent   = $isCurrent
            }
        }
    }

    return $versions
}

function Get-NodeExeLocations {
    $roots = @(
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)},
        $env:CommonProgramFiles,
        $env:APPDATA
    ) | Where-Object { $_ -and (Test-Path $_) }

    $all = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($root in $roots) {
        Get-ChildItem -Path $root -Filter node.exe -File -Recurse -Force -ErrorAction SilentlyContinue |
            ForEach-Object { [void]$all.Add($_.FullName) }
    }

    try {
        $wherePaths = (& where.exe node 2>$null) | Where-Object { $_ -and (Test-Path $_) }
        foreach ($p in $wherePaths) { [void]$all.Add($p) }
    } catch {}

    return @($all)
}

function Is-NvmPath {
    param([string]$Path)

    $nvmRoots = @(
        $env:NVM_HOME,
        $env:NVM_SYMLINK,
        (Join-Path $env:APPDATA "nvm")
    ) | Where-Object { $_ }

    foreach ($root in $nvmRoots) {
        if ($Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

# -----------------------------
# Update routines
# -----------------------------
function Update-NvmNode {
    if (-not (Get-Command nvm -ErrorAction SilentlyContinue)) {
        Write-Log "NVM not found in PATH. Skipping NVM update." "WARN"
        return
    }

    Invoke-Step "Read NVM installed versions" {
        $script:installedNvm = Get-NvmVersions
    }
    if (-not $script:installedNvm) {
        Write-Log "No NVM Node versions found."
        return
    }

    $installed = $script:installedNvm
    $current = $installed | Where-Object IsCurrent | Select-Object -First 1
    $byMajor = $installed | Group-Object Major

    foreach ($grp in $byMajor) {
        $major = [int]$grp.Name
        $source = ($grp.Group | Sort-Object Version -Descending | Select-Object -First 1).VersionText
        $cmd = "nvm install $major --reinstall-packages-from=$source"

        Invoke-Step "NVM update major $major from $source" {
            if ($PSCmdlet.ShouldProcess("NVM major $major", $cmd)) {
                [void](Invoke-CmdText -Command $cmd)
            }
        }
    }

    Invoke-Step "Refresh NVM versions after install" {
        $script:afterNvm = Get-NvmVersions
    }

    if ($current -and $script:afterNvm) {
        $newCurrent = $script:afterNvm |
            Where-Object { $_.Major -eq $current.Major } |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($newCurrent) {
            $useCmd = "nvm use $($newCurrent.VersionText)"
            Invoke-Step "Set active NVM version to $($newCurrent.VersionText)" {
                if ($PSCmdlet.ShouldProcess("Set active Node version", $useCmd)) {
                    [void](Invoke-CmdText -Command $useCmd)
                }
            }
        }
    }

    if ($CleanupOldNvmVersions -and $script:afterNvm) {
        $latestByMajor = @{}
        foreach ($grp in ($script:afterNvm | Group-Object Major)) {
            $latestByMajor[$grp.Name] = ($grp.Group | Sort-Object Version -Descending | Select-Object -First 1).VersionText
        }

        foreach ($v in $script:afterNvm) {
            $latest = $latestByMajor[[string]$v.Major]
            if ($v.VersionText -ne $latest) {
                $cmd = "nvm uninstall $($v.VersionText)"
                Invoke-Step "Remove old NVM version $($v.VersionText)" {
                    if ($PSCmdlet.ShouldProcess("NVM version $($v.VersionText)", $cmd)) {
                        [void](Invoke-CmdText -Command $cmd)
                    }
                }
            }
        }
    }
}

function Update-SystemNode {
    Invoke-Step "Locate node.exe installs" {
        $script:allNodeExe = Get-NodeExeLocations
    }

    if (-not $script:allNodeExe) {
        Write-Log "No node.exe found in scanned locations." "WARN"
        return
    }

    $systemNodeExe = $script:allNodeExe | Where-Object { -not (Is-NvmPath -Path $_) }
    if (-not $systemNodeExe) {
        Write-Log "Only NVM-based node.exe found; no system install to update."
        return
    }

    Write-Log "Non-NVM node.exe paths detected:"
    foreach ($p in $systemNodeExe) { Write-Log "  $p" }

    $wingetSucceeded = $false
    $hasWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)

    if (-not $UseWinget) {
        Write-Log "Winget usage disabled by parameter. Using MSI fallback." "WARN"
    } elseif (-not $hasWinget) {
        Write-Log "Winget not found in PATH. Using MSI fallback." "WARN"
    }

    if ($UseWinget -and $hasWinget) {
        $wingetId = if ($Channel -eq "LTS") { "OpenJS.NodeJS.LTS" } else { "OpenJS.NodeJS" }
        $script:wingetSucceeded = $false
        Invoke-Step "System update via winget ($wingetId)" {
            $cmd = "winget upgrade --id $wingetId -e --silent --accept-source-agreements --accept-package-agreements"
            if ($PSCmdlet.ShouldProcess("System Node.js via winget", $cmd)) {
                & winget upgrade --id $wingetId -e --silent --accept-source-agreements --accept-package-agreements | Out-Null
                $script:wingetSucceeded = $true
            }
        }
        $wingetSucceeded = [bool]$script:wingetSucceeded
    }

    if ($wingetSucceeded) { return }

    Invoke-Step "Get latest Node.js metadata from nodejs.org" {
        $index = Invoke-RestMethod -Uri "https://nodejs.org/dist/index.json"
        $target = if ($Channel -eq "LTS") {
            $index | Where-Object { $_.lts } | Select-Object -First 1
        } else {
            $index | Select-Object -First 1
        }
        if (-not $target) {
            throw "Could not resolve target Node.js version for channel '$Channel'."
        }
        $script:targetNode = $target
    }

    if (-not $script:targetNode) { return }

    $ver = ($script:targetNode.version -replace '^v', '')
    $msi = "node-v$ver-x64.msi"
    $url = "https://nodejs.org/dist/v$ver/$msi"
    $tmp = Join-Path $env:TEMP $msi

    Invoke-Step "Download MSI $msi" {
        if ($PSCmdlet.ShouldProcess("Download Node.js $ver", $url)) {
            Invoke-WebRequest -Uri $url -OutFile $tmp
        }
    }

    Invoke-Step "Install MSI $msi" {
        if ($PSCmdlet.ShouldProcess("Install Node.js $ver", "msiexec /i $tmp /qn /norestart")) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList @("/i", $tmp, "/qn", "/norestart") -Wait -NoNewWindow
        }
    }

    Invoke-Step "Cleanup MSI temp file" {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

# -----------------------------
# Main
# -----------------------------
try {
    Write-Log "Node update started. Log file: $script:LogFile"

    if ($UpdateNvm) {
        Invoke-Step "Update NVM-managed Node.js" { Update-NvmNode }
    } else {
        Write-Log "NVM update disabled by parameter."
    }

    if ($UpdateSystemInstall) {
        Invoke-Step "Update system Node.js install" { Update-SystemNode }
    } else {
        Write-Log "System install update disabled by parameter."
    }

    Write-Log "Node update process completed."
}
finally {
    Write-Host ""
    Write-Host "Step Summary"
    Write-Host "------------"
    if ($script:StepResults.Count -gt 0) {
        $script:StepResults | Format-Table -AutoSize | Out-String | Write-Host
    } else {
        Write-Host "No steps were recorded."
    }

    Write-Host "Log file: $script:LogFile"
    Read-Host "Press Enter to exit"
}
