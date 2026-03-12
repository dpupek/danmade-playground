# =====================================================================
# Winget updater (interactive selection)
# - Lists upgrades with numbers so the user can pick specific items
# - Runs winget only for the chosen packages
# - Still re-launches elevated to catch machine-scope installs
# =====================================================================

param(
  [switch]$MachinePhase, # internal flag when relaunching elevated
  [switch]$NonSilentPhase, # internal flag for elevated reruns without --silent
  [string]$SelectedList  # comma-separated package ids when elevated
)

$script:InstallerFailures = @()
$script:CachedUninstallEntries = $null

# Args for single-package upgrade runs. Do not include --all/--recurse (those are for bulk upgrades).
$WingetCommonArgs = @(
  "--include-unknown",
  "--silent",
  "--exact",
  "--accept-package-agreements",
  "--accept-source-agreements"
)

function Normalize-DisplayText {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $normalized = $Text.ToLowerInvariant()
  $normalized = $normalized -replace '\b\d+(?:\.\d+)+(?:[-_]\d+)?\b', ' '
  $normalized = $normalized -replace '[^a-z0-9]+', ' '
  $normalized = $normalized -replace '\s+', ' '
  return $normalized.Trim()
}

function Get-CachedUninstallEntries {
  if ($script:CachedUninstallEntries) { return $script:CachedUninstallEntries }

  function Get-OptionalPropertyValue {
    param(
      [object]$Object,
      [string]$Name
    )

    if (-not $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if (-not $prop) { return $null }
    return [string]$prop.Value
  }

  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  $script:CachedUninstallEntries = @(
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
      Where-Object {
        $_.PSObject.Properties['DisplayName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$_.PSObject.Properties['DisplayName'].Value)
      } |
      ForEach-Object {
        [pscustomobject]@{
          DisplayName      = Get-OptionalPropertyValue -Object $_ -Name 'DisplayName'
          DisplayNameNorm  = Normalize-DisplayText -Text (Get-OptionalPropertyValue -Object $_ -Name 'DisplayName')
          DisplayVersion   = Get-OptionalPropertyValue -Object $_ -Name 'DisplayVersion'
          InstallLocation  = Get-OptionalPropertyValue -Object $_ -Name 'InstallLocation'
          UninstallString  = Get-OptionalPropertyValue -Object $_ -Name 'UninstallString'
          PSPath           = [string]$_.PSPath
        }
      }
  )

  return $script:CachedUninstallEntries
}

function Get-ScopeFromEntry {
  param([object]$Entry)

  $psPath = [string]$Entry.PSPath
  $installLocation = ([string]$Entry.InstallLocation).Trim('"')
  $uninstallString = ([string]$Entry.UninstallString).Trim('"')
  $localAppData = [regex]::Escape($env:LOCALAPPDATA)
  $userProfile = [regex]::Escape($env:USERPROFILE)
  $programFiles = [regex]::Escape($env:ProgramFiles)
  $programFilesX86Path = ${env:ProgramFiles(x86)}
  $programFilesX86 = if ($programFilesX86Path) { [regex]::Escape($programFilesX86Path) } else { $null }

  if ($psPath -match 'HKEY_CURRENT_USER') { return 'User' }
  if ($psPath -match 'HKEY_LOCAL_MACHINE') { return 'Machine' }
  if ($installLocation -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($uninstallString -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($installLocation -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $installLocation -match "^$programFilesX86") { return 'Machine' }
  if ($uninstallString -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $uninstallString -match "^$programFilesX86") { return 'Machine' }

  return 'Unknown'
}

function Resolve-PackageInstallMetadata {
  param([object]$Package)

  if (-not $Package -or [string]::IsNullOrWhiteSpace($Package.Name)) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $packageName = [string]$Package.Name
  $packageNameNorm = Normalize-DisplayText -Text $packageName
  if (-not $packageNameNorm) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $bestMatch = $null
  $bestScore = -1
  foreach ($entry in @(Get-CachedUninstallEntries)) {
    if (-not $entry.DisplayNameNorm) { continue }

    $score = -1
    if ($entry.DisplayName -eq $packageName) {
      $score = 140
    } elseif ($entry.DisplayNameNorm -eq $packageNameNorm) {
      $score = 120
    } elseif ($entry.DisplayNameNorm.StartsWith($packageNameNorm)) {
      $score = 100
    } elseif ($entry.DisplayNameNorm.Contains($packageNameNorm)) {
      $score = 90
    } elseif ($packageNameNorm.Contains($entry.DisplayNameNorm)) {
      $score = 70
    }

    if ($score -lt 0) { continue }
    if ($score -gt $bestScore) {
      $bestScore = $score
      $bestMatch = $entry
    }
  }

  if (-not $bestMatch) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $scope = Get-ScopeFromEntry -Entry $bestMatch
  $note = switch ($scope) {
    'User' { 'Per-user install; keep this in the current session.' }
    'Machine' { 'Machine install; elevated retry can apply.' }
    default { $null }
  }

  return [pscustomobject]@{
    Scope = $scope
    Note  = $note
  }
}

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

function Normalize-WingetPackageId {
  param([string]$RawId)

  if ([string]::IsNullOrWhiteSpace($RawId)) { return $null }

  $candidate = $RawId.Trim()

  # Strip leading non-identifier artifacts from mojibake/truncated table output.
  $candidate = $candidate -replace '^[^A-Za-z0-9]+', ''

  # Keep the first token that looks like a winget package identifier.
  # Support both dotted IDs (e.g., Microsoft.Teams) and non-dotted store IDs.
  $match = [regex]::Match($candidate, '[A-Za-z0-9][A-Za-z0-9._+-]*')
  if ($match.Success) { return $match.Value }

  return $candidate
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

    $idRaw = Get-FirstValue -Object $Node -Names @('PackageIdentifier', 'Id')
    $id = Normalize-WingetPackageId -RawId $idRaw
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
            if ($deduped) {
              return @(
                foreach ($pkg in $deduped) {
                  $metadata = Resolve-PackageInstallMetadata -Package $pkg
                  $pkg | Add-Member -NotePropertyName InstallScope -NotePropertyValue $metadata.Scope -Force
                  $pkg | Add-Member -NotePropertyName ScopeNote -NotePropertyValue $metadata.Note -Force
                  $pkg
                }
              )
            }
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

    $headerIndex = [array]::IndexOf($lines, $header)
    $candidateLines = $lines | Select-Object -Skip ($headerIndex + 2)
    # Some winget rows collapse to single-space separators when columns are full width.
    # Accept both dotted and non-dotted package identifiers.
    $rowPattern = '^(?<Name>.+?)\s+(?<Id>[A-Za-z0-9][A-Za-z0-9._+-]*)\s+(?<Installed>\S+)\s+(?<Available>\S+)\s+(?<Source>\S+)\s*$'

    $parsed = foreach ($line in $candidateLines) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      if ($line -match '^[-\s]+$') { continue }
      if ($line -match '^\d+\s+upgrades available\.?$') { continue }
      if ($line -match '^The following packages have an upgrade available') { continue }

      if ($line -notmatch $rowPattern) { continue }
      $name = $matches['Name'].Trim()
      $id = Normalize-WingetPackageId -RawId $matches['Id']
      $installed = $matches['Installed'].Trim()
      $available = $matches['Available'].Trim()
      $source = $matches['Source'].Trim()

      if (-not $id) { continue }

      [pscustomobject]@{
        Name      = $name
        Id        = $id
        Installed = $installed
        Available = $available
        Source    = $source
      }
    }
    return @(
      foreach ($pkg in $parsed) {
        $metadata = Resolve-PackageInstallMetadata -Package $pkg
        $pkg | Add-Member -NotePropertyName InstallScope -NotePropertyValue $metadata.Scope -Force
        $pkg | Add-Member -NotePropertyName ScopeNote -NotePropertyValue $metadata.Note -Force
        $pkg
      }
    )
  } catch {
    Write-Warning "winget upgrade listing failed: $($_.Exception.Message)"
    return @()
  }
}

function Show-NumberedUpgrades($packages) {
  Write-Host "`n=== Winget upgrade candidates (including unknown versions) ===`n"
  Write-Host "Note: packages marked [User-scope] should stay in the current session; elevated machine-scope runs do not apply to them.`n" -ForegroundColor DarkYellow
  $i = 1
  foreach ($pkg in $packages) {
    $installed = if ($pkg.Installed) { $pkg.Installed } else { 'unknown' }
    $available = if ($pkg.Available) { $pkg.Available } else { 'unknown' }
    $scopeTag = switch ($pkg.InstallScope) {
      'User' { ' [User-scope]' }
      'Machine' { ' [Machine-scope]' }
      default { '' }
    }
    $line = "[{0}] {1} ({2} -> {3}) | Id: {4}{5}" -f `
            $i, $pkg.Name, $installed, $available, $pkg.Id, `
            ($(if ($pkg.Source) { " | Source: $($pkg.Source)" } else { '' }))
    if ($pkg.ScopeNote) { $line += " | Note: $($pkg.ScopeNote)" }
    Write-Host ($line + $scopeTag)
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

  if ($Failure.WingetHint) { return $Failure.WingetHint }
  if ($Failure.Stage -like 'Machine-scope*' -and $Failure.InstallScope -eq 'User') {
    return "Retry this package in the current session instead of the elevated window."
  }
  if ($Failure.Stage -eq 'Current session') {
    return "Retry in Elevated window for this package."
  }
  return "Close related apps/processes and retry."
}

function Get-WingetFailureHint {
  param(
    [int]$WingetExitCode,
    [object]$Package,
    [string]$Stage,
    [string]$LogPath
  )

  switch ($WingetExitCode) {
    -1978335189 {
      $diagText = $null
      if ($LogPath -and (Test-Path $LogPath)) {
        try { $diagText = Get-Content -Path $LogPath -Raw -ErrorAction Stop } catch { $diagText = $null }
      }

      if ($diagText -and $diagText -match 'Installer scope does not match currently installed scope:\s*(?<available>\w+)\s*!=\s*(?<installed>\w+)') {
        $availableScope = $Matches['available']
        $installedScope = $Matches['installed']
        return "Installed as $($installedScope.ToLowerInvariant())-scope, but the available winget upgrade only supports $($availableScope.ToLowerInvariant())-scope installs. Winget cannot upgrade this install in place."
      }
      if ($Package -and $Package.InstallScope -eq 'User' -and $Stage -like 'Machine-scope*') {
        return "This package is installed per-user, so the machine-scope run does not apply. Run or retry it in the current session instead."
      }
      return "A newer version exists, but winget says it does not apply to this install or to this system's requirements."
    }
    default { return $null }
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
    [string]$Stage,
    [switch]$NonSilent
  )

  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }

  foreach ($pkg in @($Packages)) {
    if (-not $pkg.Id) { continue }
    $safeId = ($pkg.Id -replace '[^A-Za-z0-9._-]','_')
    $logDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winget-update-script-logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $wingetLogPath = Join-Path -Path $logDir -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    $args = @('upgrade','--id',$pkg.Id,'--log',$wingetLogPath) + $WingetCommonArgs
    if ($NonSilent) {
      $args = @($args | Where-Object { $_ -ne '--silent' })
    }
    if ($pkg.Source) { $args += @('--source', $pkg.Source) }
    Write-Host "`n[$Stage] Running: winget $($args -join ' ')"

    & winget @args
    $wingetExitCode = $LASTEXITCODE

    $installerExitCode = Get-InstallerExitCodeFromLog -Path $wingetLogPath

    if ($wingetExitCode -ne 0) {
      $wingetHex = Convert-ToHexCode -Code $wingetExitCode
      $wingetHint = Get-WingetFailureHint -WingetExitCode $wingetExitCode -Package $pkg -Stage $Stage -LogPath $wingetLogPath
      $installerHint = Get-InstallerExitHint -InstallerExitCode $installerExitCode -OutputLines @()
      $warningText = "winget failed for $($pkg.Id) in $Stage. winget code: $wingetExitCode ($wingetHex)"
      if ($installerExitCode -ne $null) {
        $warningText += "; installer code: $installerExitCode"
      }
      if ($wingetHint) {
        $warningText += ". $wingetHint"
      } elseif ($installerHint) {
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
        WingetHint        = $wingetHint
        RetryHint         = $null
        LogPath           = $wingetLogPath
        InstallScope      = $pkg.InstallScope
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
  param(
    [string[]]$Ids,
    [switch]$NonSilent
  )

  if (-not $Ids -or $Ids.Count -eq 0) { return }

  $csv = ($Ids -join ',')
  $argList = @(
    '-NoProfile',
    '-ExecutionPolicy','Bypass',
    '-File',"`"$PSCommandPath`"",
    '-MachinePhase'
  )
  if ($NonSilent) {
    $argList += '-NonSilentPhase'
  }
  $argList += @(
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

function Prompt-RetryFailedNonSilent {
  param([string]$ScopeDescription = "this window")

  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Retry failed packages without --silent in $ScopeDescription"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not retry now")
  )
  $message = "Retry failed package upgrades without --silent? This can surface installer dialogs/prompts that were hidden during the silent run. Target: $ScopeDescription."
  $result = $Host.UI.PromptForChoice("Retry failed upgrades without --silent", $message, $choices, 0)
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
        $failedPackages = @(
          foreach ($failure in $script:InstallerFailures) {
            $selectedPackages | Where-Object { $_.Id -eq $failure.Id } | Select-Object -First 1
          }
        ) | Where-Object { $_ } | Select-Object -Unique

        if ($failedPackages.Count -gt 0 -and (Prompt-RetryFailedNonSilent -ScopeDescription "the current session")) {
          $script:InstallerFailures = @()
          Start-WingetSelected -Packages $failedPackages -Stage "Current session (non-silent retry)" -NonSilent
          if ($script:InstallerFailures.Count -gt 0) {
            Summarize-Failures -Failures $script:InstallerFailures
          } else {
            Write-Host "`nNon-silent retry completed successfully."
          }
        }

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
    $userScopedPackages = @($selectedPackages | Where-Object { $_.InstallScope -eq 'User' })
    $machineScopedPackages = @($selectedPackages | Where-Object { $_.InstallScope -ne 'User' })

    if ($userScopedPackages.Count -gt 0) {
      $userScopedIds = ($userScopedPackages | Select-Object -ExpandProperty Id) -join ', '
      Write-Host "Running user-scope packages in the current session: $userScopedIds" -ForegroundColor Yellow
      Start-WingetSelected -Packages $userScopedPackages -Stage "Current session (user-scope)"
      if ($script:InstallerFailures.Count -gt 0) {
        Summarize-Failures -Failures $script:InstallerFailures
      }
    }

    if ($machineScopedPackages.Count -gt 0) {
      Start-MachinePhase -Ids ($machineScopedPackages | Select-Object -ExpandProperty Id)
      Write-Host "Opening an elevated window to run machine-scope upgrades..."
    } else {
      Write-Host "No machine-scope packages remain after filtering out user-scope installs."
    }
    Exit-WithPause
  }

  # Phase 2 (elevated): machine-scope upgrade for the same selection
  $machineIds = if ($SelectedList) { $SelectedList -split ',' | Where-Object { $_ } } else { @() }
  if (-not $machineIds -or $machineIds.Count -eq 0) {
    Write-Warning "No package ids provided to machine phase; nothing to do."
    Exit-WithPause
  }

  $machinePackages = foreach ($id in $machineIds) { [pscustomobject]@{ Name = $id; Id = $id; Source = $null } }
  $machineStage = if ($NonSilentPhase) { "Machine-scope (non-silent retry)" } else { "Machine-scope" }
  Start-WingetSelected -Packages $machinePackages -Stage $machineStage -NonSilent:$NonSilentPhase
  if ($script:InstallerFailures.Count -gt 0) {
    Write-Host "`n$machineStage completed with failures."
    Analyze-FailuresWithCodex -Failures $script:InstallerFailures
    Summarize-Failures -Failures $script:InstallerFailures

    if (-not $NonSilentPhase) {
      $failedIds = @($script:InstallerFailures | Select-Object -ExpandProperty Id -Unique)
      if ($failedIds.Count -gt 0 -and (Prompt-RetryFailedNonSilent -ScopeDescription "this elevated window")) {
        $script:InstallerFailures = @()
        $retryPackages = foreach ($id in $failedIds) { [pscustomobject]@{ Name = $id; Id = $id; Source = $null } }
        Start-WingetSelected -Packages $retryPackages -Stage "Machine-scope (non-silent retry)" -NonSilent

        if ($script:InstallerFailures.Count -gt 0) {
          Write-Host "`nMachine-scope non-silent retry still has failures."
          Analyze-FailuresWithCodex -Failures $script:InstallerFailures
          Summarize-Failures -Failures $script:InstallerFailures
        } else {
          Write-Host "Machine-scope non-silent retry completed successfully."
        }
      }
    }
  } else {
    Write-Host "$machineStage upgrades complete."
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
