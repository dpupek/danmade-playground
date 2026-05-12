<#
.SYNOPSIS
Runs winget upgrades on one or more remote Windows computers over PowerShell remoting.

.DESCRIPTION
This script is a non-interactive companion to winget-ms-store-update-all.ps1.
It is intended for WinRM/PowerShell remoting scenarios where GUI pickers, local
UAC relaunch, and end-of-run pause prompts are not viable.

By default, the script only lists available upgrades on the target computer.
Use -All or -PackageId to actually install updates.

.NOTES
- Requires PowerShell remoting access to the target computer(s).
- The remote session should already have the privileges needed to run winget.
- Microsoft Store and per-user packages can still be constrained by the remote
  session context on the target machine.
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = "List")]
param(
    [Parameter(Mandatory, Position = 0)]
    [string[]]$ComputerName,

    [Parameter(ParameterSetName = "ById")]
    [string[]]$PackageId,

    [Parameter(ParameterSetName = "All")]
    [switch]$All,

    [pscredential]$Credential,

    [switch]$UseSSL,

    [string]$ConfigurationName,

    [int]$ThrottleLimit = 8,

    [string]$RemoteLogDirectory = "C:\ProgramData\winget-update-script-logs",

    [switch]$SkipMsStoreSourceSetup,

    [switch]$IncludeUnknown
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$remoteScript = {
    param(
        [string[]]$RequestedPackageIds,
        [bool]$UpgradeAll,
        [string]$LogDirectory,
        [bool]$SkipMsStoreSetup,
        [bool]$IncludeUnknownSwitch,
        [bool]$PreviewOnly
    )

    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"

    function New-ResultObject {
        param(
            [string]$ComputerName,
            [string]$Action,
            [string]$Status,
            [string]$Id,
            [string]$Name,
            [string]$Installed,
            [string]$Available,
            [string]$Source,
            [int]$ExitCode,
            [string]$LogPath,
            [string]$Message
        )

        [pscustomobject]@{
            ComputerName = $ComputerName
            Action       = $Action
            Status       = $Status
            Id           = $Id
            Name         = $Name
            Installed    = $Installed
            Available    = $Available
            Source       = $Source
            ExitCode     = $ExitCode
            LogPath      = $LogPath
            Message      = $Message
        }
    }

    function Normalize-WingetPackageId {
        param([string]$RawId)

        if ([string]::IsNullOrWhiteSpace($RawId)) { return $null }

        $candidate = $RawId.Trim()
        $candidate = $candidate -replace '^[^A-Za-z0-9]+', ''

        $match = [regex]::Match($candidate, '[A-Za-z0-9][A-Za-z0-9._+-]*')
        if ($match.Success) { return $match.Value }

        return $candidate
    }

    function Get-FirstValue {
        param(
            [object]$Object,
            [string[]]$Names
        )

        if (-not $Object) { return $null }

        $propNames = $Object.PSObject.Properties.Name
        foreach ($name in $Names) {
            if ($propNames -notcontains $name) { continue }

            $value = $Object.$name
            if ($null -eq $value) { continue }

            $text = [string]$value
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                return $text.Trim()
            }
        }

        return $null
    }

    function Add-WingetJsonPackages {
        param(
            [object]$Node,
            [System.Collections.Generic.List[object]]$Collector,
            [string]$InheritedSource
        )

        if ($null -eq $Node) { return }

        if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
            foreach ($child in $Node) {
                Add-WingetJsonPackages -Node $child -Collector $Collector -InheritedSource $InheritedSource
            }
            return
        }

        $nodeProps = $Node.PSObject.Properties.Name
        if (-not $nodeProps) { return }

        $sourceFromNode = $InheritedSource
        if ($nodeProps -contains "Source") {
            $candidate = Get-FirstValue -Object $Node -Names @("Source")
            if ($candidate) { $sourceFromNode = $candidate }
        } elseif (($nodeProps -contains "Name") -and ($nodeProps -contains "Packages")) {
            $candidate = Get-FirstValue -Object $Node -Names @("Name")
            if ($candidate) { $sourceFromNode = $candidate }
        } elseif ($nodeProps -contains "Details") {
            $candidate = Get-FirstValue -Object $Node.Details -Names @("Name")
            if ($candidate) { $sourceFromNode = $candidate }
        }

        $idRaw = Get-FirstValue -Object $Node -Names @("PackageIdentifier", "Id")
        $id = Normalize-WingetPackageId -RawId $idRaw
        if ($id) {
            $name = Get-FirstValue -Object $Node -Names @("PackageName", "Name")
            if (-not $name) { $name = $id }

            $installed = Get-FirstValue -Object $Node -Names @("InstalledVersion", "Version", "Installed")
            $available = Get-FirstValue -Object $Node -Names @("AvailableVersion", "Available", "LatestVersion")
            $source = Get-FirstValue -Object $Node -Names @("Source", "Repository", "Origin")
            if (-not $source) { $source = $sourceFromNode }

            if ($installed -or $available) {
                $Collector.Add([pscustomobject]@{
                    Name      = $name
                    Id        = $id
                    Installed = $installed
                    Available = $available
                    Source    = $source
                })
            }
        }

        foreach ($prop in $Node.PSObject.Properties) {
            if ($null -eq $prop.Value) { continue }
            if ($prop.Value -is [string]) { continue }
            Add-WingetJsonPackages -Node $prop.Value -Collector $Collector -InheritedSource $sourceFromNode
        }
    }

    function Get-WingetUpgradeList {
        $baseArgs = @("upgrade")
        if ($IncludeUnknownSwitch) {
            $baseArgs += "--include-unknown"
        }

        try {
            $json = & winget @($baseArgs + @("--output", "json")) 2>$null
            if ($json) {
                $jsonText = ($json -join "`n").Trim()
                if ($jsonText) {
                    $firstBraceIndex = $jsonText.IndexOfAny(@("[", "{"))
                    if ($firstBraceIndex -ge 0) {
                        $jsonBody = $jsonText.Substring($firstBraceIndex)
                        $data = $jsonBody | ConvertFrom-Json -ErrorAction Stop
                        $collected = New-Object System.Collections.Generic.List[object]
                        Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null

                        if ($collected.Count -gt 0) {
                            $seen = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
                            return @(
                                foreach ($pkg in $collected) {
                                    if (-not $pkg.Id) { continue }
                                    if ($seen.Add($pkg.Id)) { $pkg }
                                }
                            )
                        }
                    }
                }
            }
        } catch {
        }

        try {
            $output = & winget @baseArgs 2>$null
            if (-not $output) { return @() }

            $lines = $output -split "`r?`n"
            $header = $lines | Where-Object { $_ -match '^Name\s+Id\s+Version\s+Available\s+Source' } | Select-Object -First 1
            if (-not $header) { return @() }

            $headerIndex = [array]::IndexOf($lines, $header)
            $candidateLines = $lines | Select-Object -Skip ($headerIndex + 2)
            $rowPattern = '^(?<Name>.+?)\s+(?<Id>[A-Za-z0-9][A-Za-z0-9._+-]*)\s+(?<Installed>\S+)\s+(?<Available>\S+)\s+(?<Source>\S+)\s*$'

            return @(
                foreach ($line in $candidateLines) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    if ($line -match '^[-\s]+$') { continue }
                    if ($line -match '^\d+\s+upgrades available\.?$') { continue }
                    if ($line -match '^The following packages have an upgrade available') { continue }
                    if ($line -notmatch $rowPattern) { continue }

                    $id = Normalize-WingetPackageId -RawId $matches["Id"]
                    if (-not $id) { continue }

                    [pscustomobject]@{
                        Name      = $matches["Name"].Trim()
                        Id        = $id
                        Installed = $matches["Installed"].Trim()
                        Available = $matches["Available"].Trim()
                        Source    = $matches["Source"].Trim()
                    }
                }
            )
        } catch {
            return @()
        }
    }

    function Ensure-WingetMsStoreSource {
        if ($SkipMsStoreSetup) { return }

        try {
            $sources = & winget source list 2>$null
            if ($sources -notmatch "msstore") {
                & winget source add --name msstore --arg https://storeedgefd.dsx.mp.microsoft.com/v9.0 | Out-Null
            }
            & winget source update | Out-Null
        } catch {
        }
    }

    function Get-InstallerExitCodeFromLog {
        param([string]$Path)

        if (-not $Path -or -not (Test-Path $Path)) { return $null }

        try {
            $text = Get-Content -Path $Path -Raw -ErrorAction Stop
            $patterns = @(
                '(?i)(?:install|uninstall)\s+failed\s+with\s+exit\s+code:\s*(-?\d+)',
                '(?i)installer\s+return\s+code\s*[:=]\s*(-?\d+)',
                '(?i)msi(?:\s+installer)?\s+(?:return|exit)\s+code\s*[:=]\s*(-?\d+)'
            )

            foreach ($pattern in $patterns) {
                $match = [regex]::Match($text, $pattern)
                if (-not $match.Success) { continue }

                $parsed = 0
                if ([int]::TryParse($match.Groups[1].Value, [ref]$parsed)) {
                    return $parsed
                }
            }
        } catch {
        }

        return $null
    }

    $computer = $env:COMPUTERNAME

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        return New-ResultObject -ComputerName $computer -Action "Preflight" -Status "Failed" -Id $null -Name $null -Installed $null -Available $null -Source $null -ExitCode 0 -LogPath $null -Message "winget was not found on the remote computer."
    }

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    Ensure-WingetMsStoreSource

    $availableUpgrades = @(Get-WingetUpgradeList)
    if ($availableUpgrades.Count -eq 0) {
        return New-ResultObject -ComputerName $computer -Action "List" -Status "None" -Id $null -Name $null -Installed $null -Available $null -Source $null -ExitCode 0 -LogPath $null -Message "No upgrades were reported by winget."
    }

    if ($PreviewOnly) {
        return @(
            foreach ($pkg in $availableUpgrades) {
                New-ResultObject -ComputerName $computer -Action "List" -Status "Available" -Id $pkg.Id -Name $pkg.Name -Installed $pkg.Installed -Available $pkg.Available -Source $pkg.Source -ExitCode 0 -LogPath $null -Message "Upgrade available."
            }
        )
    }

    $requestedSet = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in @($RequestedPackageIds)) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        [void]$requestedSet.Add($id.Trim())
    }

    $selectedPackages = if ($UpgradeAll) {
        $availableUpgrades
    } else {
        @($availableUpgrades | Where-Object { $requestedSet.Contains($_.Id) })
    }

    if (-not $UpgradeAll -and $requestedSet.Count -gt 0) {
        $selectedIdSet = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($pkg in $selectedPackages) {
            [void]$selectedIdSet.Add($pkg.Id)
        }

        $missing = @(
            foreach ($requested in $requestedSet) {
                if (-not $selectedIdSet.Contains($requested)) { $requested }
            }
        )

        if ($missing.Count -gt 0) {
            foreach ($missingId in $missing) {
                New-ResultObject -ComputerName $computer -Action "Upgrade" -Status "Skipped" -Id $missingId -Name $null -Installed $null -Available $null -Source $null -ExitCode 0 -LogPath $null -Message "Requested package was not present in the available upgrade list."
            }
        }
    }

    if ($selectedPackages.Count -eq 0) {
        return New-ResultObject -ComputerName $computer -Action "Upgrade" -Status "Skipped" -Id $null -Name $null -Installed $null -Available $null -Source $null -ExitCode 0 -LogPath $null -Message "No matching upgrades were selected on the remote computer."
    }

    $commonArgs = @(
        "--silent",
        "--exact",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--disable-interactivity"
    )

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($pkg in $selectedPackages) {
        $safeId = ($pkg.Id -replace '[^A-Za-z0-9._-]', '_')
        $logPath = Join-Path -Path $LogDirectory -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format "yyyyMMdd-HHmmss"))
        $args = @("upgrade", "--id", $pkg.Id, "--log", $logPath) + $commonArgs
        if ($IncludeUnknownSwitch) {
            $args += "--include-unknown"
        }
        if ($pkg.Source) {
            $args += @("--source", $pkg.Source)
        }

        & winget @args
        $wingetExitCode = if ($null -ne $LASTEXITCODE) { [int]$LASTEXITCODE } else { 0 }
        $installerExitCode = Get-InstallerExitCodeFromLog -Path $logPath

        if ($wingetExitCode -eq 0) {
            $results.Add((New-ResultObject -ComputerName $computer -Action "Upgrade" -Status "Succeeded" -Id $pkg.Id -Name $pkg.Name -Installed $pkg.Installed -Available $pkg.Available -Source $pkg.Source -ExitCode 0 -LogPath $logPath -Message "Upgrade completed successfully."))
            continue
        }

        $message = "winget failed with exit code $wingetExitCode."
        if ($installerExitCode -ne $null) {
            $message += " Installer exit code: $installerExitCode."
        }

        $results.Add((New-ResultObject -ComputerName $computer -Action "Upgrade" -Status "Failed" -Id $pkg.Id -Name $pkg.Name -Installed $pkg.Installed -Available $pkg.Available -Source $pkg.Source -ExitCode $wingetExitCode -LogPath $logPath -Message $message))
    }

    return $results
}

$requestedPackageIds = @(
    foreach ($id in @($PackageId)) {
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $id.Trim()
        }
    }
)

$previewOnly = (-not $All.IsPresent -and $requestedPackageIds.Count -eq 0)
$actionDescription = if ($previewOnly) {
    "List available remote winget upgrades"
} elseif ($All) {
    "Upgrade all available remote winget packages"
} else {
    "Upgrade selected remote winget packages"
}

$targetDescription = $ComputerName -join ", "
if (-not $PSCmdlet.ShouldProcess($targetDescription, $actionDescription)) {
    return
}

$invokeParams = @{
    ComputerName = $ComputerName
    ScriptBlock  = $remoteScript
    ArgumentList = @(
        @($requestedPackageIds),
        [bool]$All,
        $RemoteLogDirectory,
        [bool]$SkipMsStoreSourceSetup,
        [bool]$IncludeUnknown,
        $previewOnly
    )
    ThrottleLimit = $ThrottleLimit
    ErrorAction   = "Stop"
}

if ($Credential) {
    $invokeParams.Credential = $Credential
}

if ($UseSSL) {
    $invokeParams.UseSSL = $true
}

if ($ConfigurationName) {
    $invokeParams.ConfigurationName = $ConfigurationName
}

Invoke-Command @invokeParams
