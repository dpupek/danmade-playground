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

  function Get-FirstValue {
    param(
      [object]$Object,
      [string[]]$Names
    )

    if (-not $Object) { return $null }
    $propNames = $Object.PSObject.Properties.Name
    foreach ($name in $Names) {
      if ($propNames -contains $name) {
        $value = $Object.$name
        if ($null -eq $value) { continue }
        $text = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($text)) { return $text.Trim() }
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
    if ($nodeProps -contains 'Source') {
      $candidate = Get-FirstValue -Object $Node -Names @('Source')
      if ($candidate) { $sourceFromNode = $candidate }
    } elseif (($nodeProps -contains 'Name') -and ($nodeProps -contains 'Packages')) {
      # In many winget JSON payloads, each source entry is { Name, Packages[] }.
      $candidate = Get-FirstValue -Object $Node -Names @('Name')
      if ($candidate) { $sourceFromNode = $candidate }
    } elseif ($nodeProps -contains 'Details') {
      $candidate = Get-FirstValue -Object $Node.Details -Names @('Name')
      if ($candidate) { $sourceFromNode = $candidate }
    }

    $id = Get-FirstValue -Object $Node -Names @('PackageIdentifier', 'Id')
    if ($id) {
      $name = Get-FirstValue -Object $Node -Names @('PackageName', 'Name')
      if (-not $name) { $name = $id }

      $installed = Get-FirstValue -Object $Node -Names @('InstalledVersion', 'Version', 'Installed')
      $available = Get-FirstValue -Object $Node -Names @('AvailableVersion', 'Available', 'LatestVersion')
      $source = Get-FirstValue -Object $Node -Names @('Source', 'Repository', 'Origin')
      if (-not $source) { $source = $sourceFromNode }

      if ($available -or $installed) {
        $Collector.Add([pscustomobject]@{
          Name       = $name
          Id         = $id
          Installed  = $installed
          Available  = $available
          Source     = $source
        })
      }
    }

    foreach ($prop in $Node.PSObject.Properties) {
      if ($null -eq $prop.Value) { continue }
      if ($prop.Value -is [string]) { continue }
      Add-WingetJsonPackages -Node $prop.Value -Collector $Collector -InheritedSource $sourceFromNode
    }
  }

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
          $collected = New-Object System.Collections.Generic.List[object]
          Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null

          if ($collected.Count -gt 0) {
            $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            $deduped = foreach ($pkg in $collected) {
              if (-not $pkg.Id) { continue }
              if ($seen.Add($pkg.Id)) { $pkg }
            }
            if ($deduped) { return $deduped }
          }
        }
      }
    }
  } catch {
    # winget versions prior to full JSON support often emit "Invalid JSON primitive"; fall back quietly.
    Write-Verbose "winget JSON parse failed; using text fallback. $_"
  }

  try {
    $output = winget upgrade --include-unknown
    $lines = $output -split "`r?`n"
    $header = $lines | Where-Object { $_ -match '^Name\s+Id\s+Version\s+Available\s+Source' } | Select-Object -First 1
    if (-not $header) { return @() }

    $idxId = $header.IndexOf('Id')
    $idxVersion = $header.IndexOf('Version', $idxId + 2)
    $idxAvailable = $header.IndexOf('Available', $idxVersion + 7)
    $idxSource = $header.IndexOf('Source', $idxAvailable + 9)
    if ($idxId -lt 0 -or $idxVersion -lt 0 -or $idxAvailable -lt 0 -or $idxSource -lt 0) { return @() }

    $headerIndex = [array]::IndexOf($lines, $header)
    $candidateLines = $lines | Select-Object -Skip ($headerIndex + 2)

    $parsed = foreach ($line in $candidateLines) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      if ($line -match '^[-\s]+$') { continue }
      if ($line -match '^\d+\s+upgrades available\.?$') { continue }
      if ($line -match '^The following packages have an upgrade available') { continue }

      if ($line.Length -lt $idxSource) { continue }
      $name = $line.Substring(0, $idxId).Trim()
      $id = $line.Substring($idxId, $idxVersion - $idxId).Trim()
      $installed = $line.Substring($idxVersion, $idxAvailable - $idxVersion).Trim()
      $available = $line.Substring($idxAvailable, $idxSource - $idxAvailable).Trim()
      $source = $line.Substring($idxSource).Trim()

      if (-not $id) { continue }

      [pscustomobject]@{
        Name      = $name
        Id        = $id
        Installed = $installed
        Available = $available
        Source    = $source
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

function Convert-ToHexCode {
  param([int]$Code)
  $unsigned = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Code), 0)
  return ('0x{0:X8}' -f $unsigned)
}

function Get-InstallerExitHint {
  param(
    [Nullable[int]]$InstallerExitCode,
    [string[]]$OutputLines
  )

  if (-not $InstallerExitCode.HasValue) { return $null }

  switch ($InstallerExitCode.Value) {
    1602 { return "Installer canceled by user." }
    1603 {
      if ($OutputLines -match '(?i)uninstall failed') {
        return "MSI 1603 during uninstall/upgrade. Usually caused by locked files, running app/processes, or insufficient privileges."
      }
      return "MSI 1603 (fatal install error). Usually caused by locked files, running app/processes, or insufficient privileges."
    }
    1618 { return "Another installation is already in progress. Wait for it to finish, then retry." }
    1638 { return "Another version is already installed. Uninstall conflicting version or retry with elevation." }
    3010 { return "Installer succeeded but requires a restart to finish." }
    default { return "Installer returned exit code $($InstallerExitCode.Value)." }
  }
}

function Get-FailureRetryHint {
  param([object]$Failure)

  if ($Failure.Stage -eq 'Current session') {
    return "Retry in Elevated window for this package."
  }
  return "Close related apps/processes and retry."
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
      $m = [regex]::Match($text, $pattern)
      if (-not $m.Success) { continue }
      $tmp = 0
      if ([int]::TryParse($m.Groups[1].Value, [ref]$tmp)) {
        return $tmp
      }
    }
  } catch {
    return $null
  }

  return $null
}

function Start-WingetSelected {
  param(
    [object[]]$Packages,
    [string]$Stage
  )

  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }

  foreach ($pkg in @($Packages)) {
    if (-not $pkg.Id) { continue }
    $safeId = ($pkg.Id -replace '[^A-Za-z0-9._-]','_')
    $logDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winget-update-script-logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $wingetLogPath = Join-Path -Path $logDir -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    $args = @('upgrade','--id',$pkg.Id,'--log',$wingetLogPath) + $WingetCommonArgs
    if ($pkg.Source) { $args += @('--source', $pkg.Source) }
    Write-Host "`n[$Stage] Running: winget $($args -join ' ')"

    & winget @args
    $wingetExitCode = $LASTEXITCODE

    $installerExitCode = Get-InstallerExitCodeFromLog -Path $wingetLogPath

    if ($wingetExitCode -ne 0) {
      $wingetHex = Convert-ToHexCode -Code $wingetExitCode
      $installerHint = Get-InstallerExitHint -InstallerExitCode $installerExitCode -OutputLines @()
      $warningText = "winget failed for $($pkg.Id) in $Stage. winget code: $wingetExitCode ($wingetHex)"
      if ($installerExitCode -ne $null) {
        $warningText += "; installer code: $installerExitCode"
      }
      if ($installerHint) {
        $warningText += ". $installerHint"
      }
      Write-Warning $warningText

      $script:InstallerFailures += [pscustomobject]@{
        Id                = $pkg.Id
        Stage             = $Stage
        ExitCode          = $wingetExitCode
        ExitCodeHex       = $wingetHex
        InstallerExitCode = $installerExitCode
        InstallerHint     = $installerHint
        RetryHint         = $null
        LogPath           = $wingetLogPath
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
    if (-not $fail.RetryHint) { $fail.RetryHint = Get-FailureRetryHint -Failure $fail }
    $reason = Summarize-LogReason -Path $fail.LogPath
    $logNote = if ($fail.LogPath) { "Log: $($fail.LogPath)" } else { "No log path provided" }
    $codeNote = "winget=$($fail.ExitCode)"
    if ($fail.ExitCodeHex) { $codeNote += " ($($fail.ExitCodeHex))" }
    if ($fail.InstallerExitCode -ne $null) { $codeNote += ", installer=$($fail.InstallerExitCode)" }

    if ($reason) {
      Write-Host "- $($fail.Id): $reason [$codeNote]. $($fail.RetryHint) ($logNote)"
    } elseif ($fail.InstallerHint) {
      Write-Host "- $($fail.Id): $($fail.InstallerHint) [$codeNote]. $($fail.RetryHint) ($logNote)"
    } else {
      Write-Host "- $($fail.Id): Upgrade failed [$codeNote]. $($fail.RetryHint) ($logNote)"
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

  $currentShellPath = $null
  try { $currentShellPath = (Get-Process -Id $PID -ErrorAction Stop).Path } catch { $currentShellPath = $null }
  if (-not $currentShellPath) {
    $currentShellPath = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh' } else { 'powershell' }
  }

  try { Start-Process -FilePath $currentShellPath -Verb runas -ArgumentList $argList | Out-Null }
  catch { Write-Warning "Elevation cancelled; machine-scope upgrades skipped." }
}

function Prompt-RetryFailedInElevated {
  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Retry failed packages in an elevated window"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not retry now")
  )
  $result = $Host.UI.PromptForChoice("Retry failed upgrades", "Retry failed package upgrades in an elevated window?", $choices, 0)
  return ($result -eq 0)
}

function Exit-WithPause {
  param(
    [int]$Code = 0,
    [string]$Prompt = "Press Enter to close this window"
  )

  $script:RequestedExitCode = $Code
  $script:PausePrompt = $Prompt
  throw [System.OperationCanceledException]::new("__WINGET_SCRIPT_EXIT__")
}

# ------------------ CONTROL FLOW ------------------

$script:RequestedExitCode = 0
$script:PausePrompt = if ($MachinePhase) { "Press Enter to close this elevated window" } else { "Press Enter to close this window" }

try {
  Ensure-WingetMsStoreSource

  if (-not $MachinePhase) {
    $upgrades = Get-WingetUpgradeList
    if (-not $upgrades -or $upgrades.Count -eq 0) {
      Write-Host "No winget upgrades found."
      Exit-WithPause
    }

    Show-NumberedUpgrades -packages $upgrades
    $selectedIndexes = Prompt-Selection -Count $upgrades.Count
    if (-not $selectedIndexes -or $selectedIndexes.Count -eq 0) {
      Write-Host "No selection made. Exiting without changes."
      Exit-WithPause
    }

    $selectedPackages = foreach ($idx in $selectedIndexes) { $upgrades[$idx - 1] }
    $runMode = Prompt-RunMode
    if ($runMode -eq 0) {
      Start-WingetSelected -Packages $selectedPackages -Stage "Current session"
      if ($script:InstallerFailures.Count -gt 0) {
        Summarize-Failures -Failures $script:InstallerFailures
        $failedIds = @($script:InstallerFailures | Select-Object -ExpandProperty Id -Unique)
        if ($failedIds.Count -gt 0 -and (Prompt-RetryFailedInElevated)) {
          Start-MachinePhase -Ids $failedIds
          Write-Host "Opening an elevated window to retry failed upgrades..."
        }
      } else {
        Write-Host "`nAll selected upgrades completed successfully."
      }
      Exit-WithPause
    }

    # Relaunch elevated to run the same selection as admin.
    Start-MachinePhase -Ids ($selectedPackages | Select-Object -ExpandProperty Id)
    Write-Host "Opening an elevated window to run the selected upgrades..."
    Exit-WithPause
  }

  # Phase 2 (elevated): machine-scope upgrade for the same selection
  $machineIds = if ($SelectedList) { $SelectedList -split ',' | Where-Object { $_ } } else { @() }
  if (-not $machineIds -or $machineIds.Count -eq 0) {
    Write-Warning "No package ids provided to machine phase; nothing to do."
    Exit-WithPause
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

  Exit-WithPause -Prompt "Press Enter to close this elevated window"
} catch [System.OperationCanceledException] {
  if ($_.Exception.Message -ne "__WINGET_SCRIPT_EXIT__") {
    $script:RequestedExitCode = 1
    throw
  }
} catch {
  $script:RequestedExitCode = 1
  Write-Warning "Script failed unexpectedly: $($_.Exception.Message)"
} finally {
  Read-Host $script:PausePrompt | Out-Null
}

exit $script:RequestedExitCode
