# =====================================================================
# Winget updater (interactive selection)
# - Lists upgrades with numbers so the user can pick specific items
# - Runs winget only for the chosen packages
# - Still re-launches elevated to catch machine-scope installs
# =====================================================================

param(
  [switch]$MachinePhase, # internal flag when relaunching elevated
  [string]$SelectedList  # comma-separated package ids when elevated
)

$script:InstallerFailures = @()

# Args for single-package upgrade runs. Do not include --all/--recurse (those are for bulk upgrades).
$WingetCommonArgs = @(
  "--include-unknown",
  "--silent",
  "--exact",
  "--accept-package-agreements",
  "--accept-source-agreements"
)

function Ensure-WingetMsStoreSource {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Warning "winget not found."
    return
  }
  try {
    $sources = winget source list 2>$null
    if ($sources -notmatch 'msstore') {
      Write-Host "Adding winget msstore source..."
      winget source add --name msstore --arg https://storeedgefd.dsx.mp.microsoft.com/v9.0 | Out-Null
    }
    winget source update | Out-Null
  } catch {
    Write-Warning "Unable to verify/update winget sources: $($_.Exception.Message)"
  }
}

function Get-WingetUpgradeList {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return @() }

  # Prefer JSON for reliable parsing; fall back to table parsing.
  try {
    $json = winget upgrade --include-unknown --output json 2>$null
    if ($json) {
      $jsonText = ($json -join "`n").Trim()
      if ($jsonText) {
        $firstBraceIndex = $jsonText.IndexOfAny(@('[', '{'))
        if ($firstBraceIndex -ge 0) {
          $jsonBody = $jsonText.Substring($firstBraceIndex)
          $data = $jsonBody | ConvertFrom-Json -ErrorAction Stop
          $parsed = foreach ($item in $data) {
            $id = if ($item.PackageIdentifier) { $item.PackageIdentifier } elseif ($item.Id) { $item.Id } else { $null }
            if (-not $id) { continue }
            $installed = if ($item.Version) { $item.Version } elseif ($item.InstalledVersion) { $item.InstalledVersion } else { $null }
            $available = if ($item.Available) { $item.Available } elseif ($item.AvailableVersion) { $item.AvailableVersion } else { $null }
            [pscustomobject]@{
              Name       = $item.Name
              Id         = $id
              Installed  = $installed
              Available  = $available
              Source     = $item.Source
            }
          }
          if ($parsed) { return $parsed }
        }
      }
    }
  } catch {
    # winget versions prior to full JSON support often emit "Invalid JSON primitive"; fall back quietly.
    Write-Verbose "winget JSON parse failed; using text fallback. $_"
  }

  try {
    $output = winget upgrade --include-unknown
    $lines = $output -split "`r?`n" | Where-Object { $_ -and $_ -notmatch '^Name\s+Id' -and $_ -notmatch '^[- ]+$' }
    $parsed = foreach ($line in $lines) {
      $parts = $line -split '\s{2,}'
      if ($parts.Count -lt 4) { continue }
      [pscustomobject]@{
        Name      = $parts[0]
        Id        = $parts[1]
        Installed = $parts[2]
        Available = $parts[3]
        Source    = if ($parts.Count -ge 5) { $parts[4] } else { $null }
      }
    }
    return $parsed
  } catch {
    Write-Warning "winget upgrade listing failed: $($_.Exception.Message)"
    return @()
  }
}

function Show-NumberedUpgrades($packages) {
  Write-Host "`n=== Winget upgrade candidates (including unknown versions) ===`n"
  $i = 1
  foreach ($pkg in $packages) {
    $installed = if ($pkg.Installed) { $pkg.Installed } else { 'unknown' }
    $available = if ($pkg.Available) { $pkg.Available } else { 'unknown' }
    $line = "[{0}] {1} ({2} -> {3}) | Id: {4}{5}" -f `
            $i, $pkg.Name, $installed, $available, $pkg.Id, `
            ($(if ($pkg.Source) { " | Source: $($pkg.Source)" } else { '' }))
    Write-Host $line
    $i++
  }
}

function Prompt-Selection([int]$Count) {
  $input = Read-Host "Enter numbers (e.g., 1,3,5 or 2-4,7) or 'all' to upgrade everything; Enter to skip"
  if ([string]::IsNullOrWhiteSpace($input)) { return @() }

  $tokens = @()
  foreach ($t in ($input -split '[,\s]+')) { if ($t) { $tokens += [string]$t } }

  if ($tokens.Count -eq 1 -and ([string]$tokens[0]).ToLowerInvariant() -eq 'all') {
    return 1..$Count
  }

  $selected = New-Object System.Collections.Generic.List[int]

  foreach ($token in $tokens) {
    if ($token -match '^([0-9]+)-([0-9]+)$') {
      $start = [int]$matches[1]
      $end   = [int]$matches[2]
      if ($end -lt $start) { Write-Warning "Ignoring reversed range: $token"; continue }
      foreach ($n in $start..$end) {
        if ($n -lt 1 -or $n -gt $Count) { Write-Warning "Ignoring out-of-range selection: $n"; continue }
        if (-not $selected.Contains($n)) { $selected.Add($n) }
      }
      continue
    }

    $number = 0
    if (-not [int]::TryParse($token, [ref]$number)) {
      Write-Warning "Ignoring non-numeric entry: '$token'"
      continue
    }
    if ($number -lt 1 -or $number -gt $Count) {
      Write-Warning "Ignoring out-of-range selection: $number"
      continue
    }
    if (-not $selected.Contains($number)) { $selected.Add($number) }
  }

  return $selected
}

function Prompt-RunMode {
  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Current session (no elevation)", "Run upgrades here without elevation"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Elevated window", "Launch a UAC prompt and run selected upgrades as admin")
  )
  return $Host.UI.PromptForChoice("Run location", "Where should the selected upgrades run?", $choices, 0)
}

function Start-WingetSelected {
  param(
    [System.Collections.IEnumerable]$Packages,
    [string]$Stage
  )

  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }

  foreach ($pkg in $Packages) {
    if (-not $pkg.Id) { continue }
    $args = @('upgrade','--id',$pkg.Id) + $WingetCommonArgs
    if ($pkg.Source) { $args += @('--source', $pkg.Source) }
    Write-Host "`n[$Stage] Running: winget $($args -join ' ')"

    # Run winget via Start-Process with redirected output to avoid spinner artifacts when reprinting.
    $tmpOut = [System.IO.Path]::GetTempFileName()
    $tmpErr = [System.IO.Path]::GetTempFileName()
    $quotedArgs = ($args | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' '
    $proc = Start-Process -FilePath "winget" -ArgumentList $quotedArgs -NoNewWindow -Wait -PassThru `
                          -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr

    $output = @()
    $output += Get-Content -Path $tmpOut -ErrorAction SilentlyContinue
    $output += Get-Content -Path $tmpErr -ErrorAction SilentlyContinue
    Remove-Item -Path $tmpOut,$tmpErr -ErrorAction SilentlyContinue

    if ($output) {
      foreach ($line in $output) {
        if ($line -match '^\s*[-\\|/]\s*$') { continue } # drop spinner frames
        Write-Host $line
      }
    }

    $logPath = $null
    $logMatch = $output | Select-String -Pattern 'Installer log is available at:\s*(.+)$' | Select-Object -Last 1
    if ($logMatch) {
      $logPath = $logMatch.Matches[0].Groups[1].Value.Trim()
    }

    if ($proc.ExitCode -ne 0) {
      Write-Warning "winget returned exit code $($proc.ExitCode) for $($pkg.Id) in $Stage"
      $script:InstallerFailures += [pscustomobject]@{
        Id       = $pkg.Id
        Stage    = $Stage
        ExitCode = $proc.ExitCode
        LogPath  = $logPath
      }
    }
  }
}

function Summarize-LogReason {
  param([string]$Path)
  if (-not $Path -or -not (Test-Path $Path)) { return $null }
  try {
    $text = Get-Content -Path $Path -Raw -ErrorAction Stop

    if ($text -match 'The following process\(es\) use Git for Windows:\s*(?<procs>[^\r\n]+)') {
      $procs = $Matches['procs'].Trim()
      return "Git installer blocked because running processes: $procs. Close them and retry."
    }
    if ($text -match 'Some applications could not be shut down') {
      return "Installer could not close running applications (Restart Manager). Close WinMerge or related processes and retry."
    }
    if ($text -match 'User canceled the installation process') {
      return "Installer was canceled (likely after a prompt). Rerun and allow it to continue."
    }
    return $null
  } catch { return $null }
}

function Summarize-Failures {
  param([System.Collections.IEnumerable]$Failures)
  if (-not $Failures -or $Failures.Count -eq 0) { return }

  Write-Host "" # blank line
  Write-Host "Failure summary:" -ForegroundColor Yellow
  foreach ($fail in $Failures) {
    $reason = Summarize-LogReason -Path $fail.LogPath
    $logNote = if ($fail.LogPath) { "Log: $($fail.LogPath)" } else { "No log path provided" }
    if ($reason) {
      Write-Host "- $($fail.Id): $reason ($logNote)"
    } else {
      Write-Host "- $($fail.Id): Exit code $($fail.ExitCode). $logNote"
    }
  }
}

function Analyze-FailuresWithCodex {
  param([System.Collections.IEnumerable]$Failures)

  if (-not $Failures -or $Failures.Count -eq 0) { return }

  $codex = Get-Command codex -ErrorAction SilentlyContinue
  if (-not $codex) {
    Write-Host "Codex executable not found; skipping auto-analysis. Logs are listed above if available."
    return
  }

  foreach ($fail in $Failures) {
    if (-not $fail.LogPath -or -not (Test-Path $fail.LogPath)) {
      Write-Host "No log file found for $($fail.Id); skipping Codex analysis."
      continue
    }
    try {
      $logText = Get-Content -Path $fail.LogPath -Raw -ErrorAction Stop
      $prompt = "Summarize in plain language why the installer failed and suggest the next step to fix it. Log follows:\n\n$logText"
      Write-Host "\n--- Codex analysis for $($fail.Id) ($($fail.Stage)) ---"
      $analysis = $prompt | & $codex.Path 2>&1
      if ($analysis) { $analysis | ForEach-Object { Write-Host $_ } }
    } catch {
      Write-Warning "Codex analysis failed for $($fail.Id): $($_.Exception.Message)"
    }
  }
}

function Start-MachinePhase {
  param([string[]]$Ids)

  if (-not $Ids -or $Ids.Count -eq 0) { return }

  $csv = ($Ids -join ',')
  $argList = @(
    '-NoProfile',
    '-ExecutionPolicy','Bypass',
    '-File',"`"$PSCommandPath`"",
    '-MachinePhase',
    '-SelectedList',"`"$csv`""
  )

  try { Start-Process "PowerShell" -Verb runas -ArgumentList $argList | Out-Null }
  catch { Write-Warning "Elevation cancelled; machine-scope upgrades skipped." }
}

# ------------------ CONTROL FLOW ------------------

Ensure-WingetMsStoreSource

if (-not $MachinePhase) {
  $upgrades = Get-WingetUpgradeList
  if (-not $upgrades -or $upgrades.Count -eq 0) {
    Write-Host "No winget upgrades found."
    exit 0
  }

  Show-NumberedUpgrades -packages $upgrades
  $selectedIndexes = Prompt-Selection -Count $upgrades.Count
  if (-not $selectedIndexes -or $selectedIndexes.Count -eq 0) {
    Write-Host "No selection made. Exiting without changes."
    exit 0
  }

  $selectedPackages = foreach ($idx in $selectedIndexes) { $upgrades[$idx - 1] }
  $runMode = Prompt-RunMode
  if ($runMode -eq 0) {
    Start-WingetSelected -Packages $selectedPackages -Stage "Current session"
    if ($script:InstallerFailures.Count -gt 0) { Summarize-Failures -Failures $script:InstallerFailures }
    exit 0
  }

  # Relaunch elevated to run the same selection as admin.
  Start-MachinePhase -Ids ($selectedPackages | Select-Object -ExpandProperty Id)
  Write-Host "Opening an elevated window to run the selected upgrades..."
  exit 0
}

# Phase 2 (elevated): machine-scope upgrade for the same selection
$machineIds = if ($SelectedList) { $SelectedList -split ',' | Where-Object { $_ } } else { @() }
if (-not $machineIds -or $machineIds.Count -eq 0) {
  Write-Warning "No package ids provided to machine phase; nothing to do."
  exit 0
}

$machinePackages = foreach ($id in $machineIds) { [pscustomobject]@{ Name = $id; Id = $id; Source = $null } }
Start-WingetSelected -Packages $machinePackages -Stage "Machine-scope"
if ($script:InstallerFailures.Count -gt 0) {
  Write-Host "\nMachine-scope completed with failures."
  Analyze-FailuresWithCodex -Failures $script:InstallerFailures
  Summarize-Failures -Failures $script:InstallerFailures
} else {
  Write-Host "Machine-scope upgrades complete."
}

Read-Host "Press Enter to close this elevated window"
