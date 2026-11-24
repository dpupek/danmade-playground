<#
.SYNOPSIS
Runs unattended winget package upgrades for domain-managed Windows endpoints.

.DESCRIPTION
Danmade Patch Agent is intended for Group Policy scheduled task deployment.
It requires a valid JSON policy, emits Wazuh-friendly Event Log and JSONL records, and performs only
bounded noninteractive recovery actions.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$PolicyPath,

  [ValidateSet('Machine', 'User')]
  [string]$Mode = 'Machine',

  [string]$RunId,

  [string]$LogRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($LogRoot)) {
  $LogRoot = if ($Mode -eq 'User' -and -not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Join-Path -Path $env:LOCALAPPDATA -ChildPath 'DanmadePatchAgent'
  } else {
    'C:\ProgramData\DanmadePatchAgent'
  }
}

$script:SchemaVersion = '1.0'
$script:EventSource = 'DanmadePatchAgent'
$script:EventLogName = 'Application'
$script:RunId = if ([string]::IsNullOrWhiteSpace($RunId)) {
  '{0}-{1}-{2}-{3}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $env:COMPUTERNAME, $Mode, ([guid]::NewGuid().ToString('N'))
} else {
  $RunId
}
$script:Policy = $null
$script:JsonlPath = $null
$script:WingetCommand = $null
$script:CachedUninstallEntries = $null
$script:CachedUserAppxPackages = $null
$script:SelectionCounts = $null
$script:ApplicabilityDeferredCount = 0
$script:RunDeadline = $null
$script:ActiveRunStatePath = $null
$script:InapplicableUpgradeStatePath = $null
$script:RunLock = $null

function New-DefaultPolicy {
  [pscustomobject]@{
    enabled           = $true
    includeUnknown    = $true
    allowedPackageIds = @()
    blockedPackageIds = @()
    packageRules      = @()
    maxRetries        = 2
    retryableInstallerExitCodes = @(1618)
    maxRunMinutes     = 120
    inapplicableUpgradeDeferralDays = 7
    maintenanceWindow = [pscustomobject]@{
      enabled = $false
      start   = '02:00'
      end     = '04:00'
    }
    rebootPolicy      = 'ReportOnly'
    wazuhReporting    = [pscustomobject]@{
      eventLog = $true
      jsonl    = $true
    }
    wingetSourceRepair = $true
    logRetentionDays   = 30
  }
}

function Get-PropertyValue {
  param(
    [object]$Object,
    [string]$Name,
    [object]$Default = $null
  )

  if ($null -eq $Object) { return $Default }
  $property = $Object.PSObject.Properties[$Name]
  if (-not $property) { return $Default }
  if ($null -eq $property.Value) { return $Default }
  return $property.Value
}

function Convert-ToStringArray {
  param([object]$Value)

  if ($null -eq $Value) { return @() }
  return @(
    foreach ($item in @($Value)) {
      if (-not [string]::IsNullOrWhiteSpace([string]$item)) {
        ([string]$item).Trim()
      }
    }
  )
}

function Convert-ToIntArray {
  param([object]$Value)

  if ($null -eq $Value) { return @() }
  return @(
    foreach ($item in @($Value)) {
      try { [int]$item } catch { }
    }
  )
}

function Convert-ToPackageRules {
  param([object]$Value)

  if ($null -eq $Value) { return @() }
  return @(
    foreach ($rule in @($Value)) {
      $id = [string](Get-PropertyValue -Object $rule -Name 'id' -Default '')
      if ([string]::IsNullOrWhiteSpace($id)) { continue }

      $mode = [string](Get-PropertyValue -Object $rule -Name 'mode' -Default 'Auto')
      if ($mode -notin @('Auto', 'ManualOnly', 'QuiescenceRequired', 'NativeUpdater')) { throw 'Invalid package rule mode.' }

      $installerType = ([string](Get-PropertyValue -Object $rule -Name 'installerType' -Default '')).Trim().ToLowerInvariant()
      if ($installerType -and $installerType -notin @('wix', 'msi', 'msix', 'exe', 'inno', 'nullsoft', 'burn')) { $installerType = '' }
      $families = @(Convert-ToStringArray -Value (Get-PropertyValue $rule 'appxPackageFamilyNames' @()))
      foreach ($family in $families) {
        if ($family -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]*_[a-z0-9]{13}$') { throw 'AppX package family names must be exact identities, including publisher ID.' }
      }

      [pscustomobject]@{
        id                = $id.Trim()
        mode              = $mode
        allowMajorVersion = [bool](Get-PropertyValue -Object $rule -Name 'allowMajorVersion' -Default $false)
        processNames      = Convert-ToStringArray -Value (Get-PropertyValue -Object $rule -Name 'processNames' -Default @())
        installerType     = $installerType
        releaseHolds      = @(Get-PropertyValue -Object $rule -Name 'releaseHolds' -Default @())
        appxPackageFamilyNames = $families
      }
    }
  )
}

function Merge-Policy {
  param([object]$LoadedPolicy)

  $default = New-DefaultPolicy
  if ($null -eq $LoadedPolicy) { return $default }

  $maintenance = Get-PropertyValue -Object $LoadedPolicy -Name 'maintenanceWindow' -Default $default.maintenanceWindow
  $reporting = Get-PropertyValue -Object $LoadedPolicy -Name 'wazuhReporting' -Default $default.wazuhReporting

  $maxRetries = [int](Get-PropertyValue -Object $LoadedPolicy -Name 'maxRetries' -Default $default.maxRetries)
  if ($maxRetries -lt 0) { $maxRetries = 0 }
  if ($maxRetries -gt 5) { $maxRetries = 5 }

  $retention = [int](Get-PropertyValue -Object $LoadedPolicy -Name 'logRetentionDays' -Default $default.logRetentionDays)
  if ($retention -lt 1) { $retention = 1 }

  $maxRunMinutes = [int](Get-PropertyValue -Object $LoadedPolicy -Name 'maxRunMinutes' -Default $default.maxRunMinutes)
  if ($maxRunMinutes -lt 15) { $maxRunMinutes = 15 }
  if ($maxRunMinutes -gt 240) { $maxRunMinutes = 240 }

  $inapplicableUpgradeDeferralDays = [int](Get-PropertyValue -Object $LoadedPolicy -Name 'inapplicableUpgradeDeferralDays' -Default $default.inapplicableUpgradeDeferralDays)
  if ($inapplicableUpgradeDeferralDays -lt 1) { $inapplicableUpgradeDeferralDays = 1 }
  if ($inapplicableUpgradeDeferralDays -gt 30) { $inapplicableUpgradeDeferralDays = 30 }

  return [pscustomobject]@{
    enabled           = [bool](Get-PropertyValue -Object $LoadedPolicy -Name 'enabled' -Default $default.enabled)
    includeUnknown    = [bool](Get-PropertyValue -Object $LoadedPolicy -Name 'includeUnknown' -Default $default.includeUnknown)
    allowedPackageIds = Convert-ToStringArray -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'allowedPackageIds' -Default @())
    blockedPackageIds = Convert-ToStringArray -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'blockedPackageIds' -Default @())
    packageRules      = Convert-ToPackageRules -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'packageRules' -Default @())
    maxRetries        = $maxRetries
    retryableInstallerExitCodes = Convert-ToIntArray -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'retryableInstallerExitCodes' -Default $default.retryableInstallerExitCodes)
    maxRunMinutes     = $maxRunMinutes
    inapplicableUpgradeDeferralDays = $inapplicableUpgradeDeferralDays
    maintenanceWindow = [pscustomobject]@{
      enabled = [bool](Get-PropertyValue -Object $maintenance -Name 'enabled' -Default $default.maintenanceWindow.enabled)
      start   = [string](Get-PropertyValue -Object $maintenance -Name 'start' -Default $default.maintenanceWindow.start)
      end     = [string](Get-PropertyValue -Object $maintenance -Name 'end' -Default $default.maintenanceWindow.end)
      machineOnly = [bool](Get-PropertyValue -Object $maintenance -Name 'machineOnly' -Default $false)
    }
    rebootPolicy      = [string](Get-PropertyValue -Object $LoadedPolicy -Name 'rebootPolicy' -Default $default.rebootPolicy)
    wazuhReporting    = [pscustomobject]@{
      eventLog = [bool](Get-PropertyValue -Object $reporting -Name 'eventLog' -Default $default.wazuhReporting.eventLog)
      jsonl    = [bool](Get-PropertyValue -Object $reporting -Name 'jsonl' -Default $default.wazuhReporting.jsonl)
    }
    wingetSourceRepair = [bool](Get-PropertyValue -Object $LoadedPolicy -Name 'wingetSourceRepair' -Default $default.wingetSourceRepair)
    logRetentionDays   = $retention
  }
}

function Resolve-PolicyPath {
  if (-not [string]::IsNullOrWhiteSpace($PolicyPath)) { return $PolicyPath }

  $localPolicyPath = Join-Path -Path $LogRoot -ChildPath 'danmade-patch-agent.policy.json'
  if (Test-Path -LiteralPath $localPolicyPath) { return $localPolicyPath }

  return $null
}

function Import-AgentPolicy {
  $resolvedPath = Resolve-PolicyPath
  if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
    throw 'A policy file is required. No packages will be upgraded without policy.'
  }

  try {
    $json = Get-Content -LiteralPath $resolvedPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($json)) {
      throw "Policy file is empty: $resolvedPath"
    }
    $loaded = $json | ConvertFrom-Json -ErrorAction Stop
    foreach ($name in @('enabled', 'includeUnknown')) {
      if ((Get-PropertyValue -Object $loaded -Name $name) -isnot [bool]) { throw "Policy requires Boolean $name." }
    }
    foreach ($name in @('allowedPackageIds', 'blockedPackageIds', 'packageRules')) {
      if ($null -eq $loaded.PSObject.Properties[$name] -or $loaded.$name -isnot [array]) { throw "Policy requires array $name." }
    }
    foreach ($rule in @($loaded.packageRules)) {
      if (-not $rule.id -or $rule.mode -notin @('Auto', 'ManualOnly', 'QuiescenceRequired', 'NativeUpdater')) { throw 'Invalid package rule.' }
      if ($rule.PSObject.Properties['appxPackageFamilyNames'] -and $rule.appxPackageFamilyNames -isnot [array]) { throw 'AppX package families must be an array.' }
      foreach ($hold in @(Get-PropertyValue -Object $rule -Name 'releaseHolds' -Default @())) {
        $until = [datetime]::MinValue
        if (-not $hold.version -or $hold.scope -notin @('machine','user') -or $hold.installerSha256 -notmatch '^[a-fA-F0-9]{64}$' -or -not [datetime]::TryParse([string]$hold.until, [ref]$until) -or -not $hold.reason) { throw 'Invalid release hold.' }
      }
      $installerType = Get-PropertyValue -Object $rule -Name 'installerType' -Default ''
      if ($installerType -and $installerType -notin @('wix','msi','msix','exe','inno','nullsoft','burn')) { throw 'Invalid installer-type restriction.' }
      if ($rule.mode -eq 'QuiescenceRequired' -and @(Get-PropertyValue -Object $rule -Name 'processNames' -Default @()).Count -eq 0) { throw 'Quiescence rule requires process names.' }
    }
    foreach ($name in @('maxRetries', 'maxRunMinutes')) {
      $value = Get-PropertyValue -Object $loaded -Name $name
      if ($value -isnot [int] -and $value -isnot [long]) { throw "Policy requires integer $name." }
    }
    if ($loaded.maxRetries -lt 0 -or $loaded.maxRetries -gt 5 -or $loaded.maxRunMinutes -lt 15 -or $loaded.maxRunMinutes -gt 240) { throw 'Policy retry or duration limit is invalid.' }
    return Merge-Policy -LoadedPolicy $loaded
  } catch {
    throw "Policy unavailable or invalid: $($_.Exception.Message)"
  }
}

function Convert-ToHexCode {
  param([int]$Code)
  $unsigned = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Code), 0)
  return ('0x{0:X8}' -f $unsigned)
}

function Get-WingetScopeArgument {
  if ($Mode -eq 'User') { return 'user' }
  return 'machine'
}

function Get-RemainingRunSeconds {
  if ($null -eq $script:RunDeadline) { return $null }
  return [math]::Floor(($script:RunDeadline - (Get-Date)).TotalSeconds)
}

function Join-ProcessArguments {
  param([string[]]$Arguments)

  return (@(
    foreach ($argument in @($Arguments)) {
      if ($null -eq $argument) { continue }
      $text = [string]$argument
      if ($text -match '[\s"]') {
        '"' + ($text -replace '"', '\"') + '"'
      } else {
        $text
      }
    }
  ) -join ' ')
}

function Invoke-WingetProcess {
  param(
    [string[]]$Arguments,
    [int]$TimeoutSeconds = 180
  )

  $stdoutPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("danmade-winget-{0}.out" -f ([guid]::NewGuid()))
  $stderrPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("danmade-winget-{0}.err" -f ([guid]::NewGuid()))
  $argumentList = Join-ProcessArguments -Arguments $Arguments

  $remainingSeconds = Get-RemainingRunSeconds
  if ($null -ne $remainingSeconds -and $remainingSeconds -le 0) {
    return [pscustomobject]@{
      ExitCode = -1
      TimedOut = $true
      StdOut   = ''
      StdErr   = 'Agent run deadline exceeded before winget could start.'
    }
  }
  if ($null -ne $remainingSeconds) {
    $TimeoutSeconds = [math]::Min($TimeoutSeconds, [math]::Max(1, [int]$remainingSeconds))
  }

  $process = New-Object System.Diagnostics.Process
  $stdoutStream = $null
  $stderrStream = $null
  try {
    # Own the native handle and copy streams directly; Start-Process loses ExitCode on PowerShell 5.1.
    $process.StartInfo.FileName = $script:WingetCommand
    $process.StartInfo.Arguments = $argumentList
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.CreateNoWindow = $true
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    $stdoutStream = [IO.File]::Open($stdoutPath, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $stderrStream = [IO.File]::Open($stderrPath, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    if (-not $process.Start()) { throw 'Native process could not start.' }
    $stdoutCopy = $process.StandardOutput.BaseStream.CopyToAsync($stdoutStream)
    $stderrCopy = $process.StandardError.BaseStream.CopyToAsync($stderrStream)
    $timedOut = -not $process.WaitForExit($TimeoutSeconds * 1000)
    if ($timedOut) { try { $process.Kill() } catch { } }
    # Descendants can inherit pipes. Bound draining too, retaining partial output on timeout.
    $drained = [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($stdoutCopy, $stderrCopy), 5000)
    $stdoutStream.Flush()
    $stderrStream.Flush()
    if ($timedOut -or -not $drained) {
      return [pscustomobject]@{ ExitCode=-1; TimedOut=$true; StdOut=(Read-ProcessText -Path $stdoutPath); StdErr=((Read-ProcessText -Path $stderrPath) + "`nNative process or output capture timed out.") }
    }
    $exitCode = $process.ExitCode
    if ($null -eq $exitCode) { throw 'Native process exit code was unavailable; success cannot be inferred.' }
    return [pscustomobject]@{
      ExitCode = [int]$exitCode
      TimedOut = $false
      StdOut   = Read-ProcessText -Path $stdoutPath
      StdErr   = Read-ProcessText -Path $stderrPath
    }
  } finally {
    $process.Dispose()
    if ($null -ne $stdoutStream) { $stdoutStream.Dispose() }
    if ($null -ne $stderrStream) { $stderrStream.Dispose() }
    Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
  }
}

function Read-ProcessText {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return '' }
  # Get-Content attaches provider metadata; do not serialize it as command output.
  $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
  if ($null -eq $text) { return '' }
  return [string]::new($text.ToCharArray())
}

function Normalize-DisplayText {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $normalized = $Text.ToLowerInvariant()
  $normalized = $normalized -replace '\b\d+(?:\.\d+)+(?:[-_]\d+)?\b', ' '
  $normalized = $normalized -replace '[^a-z0-9]+', ' '
  $normalized = $normalized -replace '\s+', ' '
  return $normalized.Trim()
}

function Ensure-Directory {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return }
  if (Test-Path -LiteralPath $Path) { return }
  if ($PSCmdlet.ShouldProcess($Path, 'Create directory')) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Initialize-AgentStorage {
  Ensure-Directory -Path $LogRoot
  Ensure-Directory -Path (Join-Path -Path $LogRoot -ChildPath 'Logs')
  Ensure-Directory -Path (Join-Path -Path $LogRoot -ChildPath 'Events')
  Ensure-Directory -Path (Join-Path -Path $LogRoot -ChildPath 'State')
  Ensure-Directory -Path (Join-Path -Path $LogRoot -ChildPath 'Inventory')
  $script:JsonlPath = Join-Path -Path $LogRoot -ChildPath 'Events\patch-agent.jsonl'
  $script:ActiveRunStatePath = Join-Path -Path $LogRoot -ChildPath 'State\active-run.json'
  $script:InapplicableUpgradeStatePath = Join-Path -Path $LogRoot -ChildPath 'State\inapplicable-upgrades.json'
}

function Write-WingetInventorySnapshot {
  param([object[]]$Packages)

  if ($WhatIfPreference) { return }
  $snapshotPath = Join-Path -Path $LogRoot -ChildPath ("Inventory\winget-upgrades-{0}.json" -f $Mode.ToLowerInvariant())
  [ordered]@{
    schemaVersion = '1.0'
    timestamp     = (Get-Date).ToString('o')
    computerName  = $env:COMPUTERNAME
    mode          = $Mode
    packages      = @($Packages | Select-Object Name,Id,Installed,Available,Source,InstallScope)
  } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $snapshotPath -Encoding UTF8
}

function Get-NotificationStatePath {
  return (Join-Path -Path $LogRoot -ChildPath 'State\pending-notification.json')
}

function Get-RestartVerificationStatePath {
  return (Join-Path -Path $LogRoot -ChildPath 'State\pending-restart-verification.json')
}

function Get-ApplicationRestartStatePath {
  return (Join-Path -Path $LogRoot -ChildPath 'State\pending-application-restart.json')
}

function Get-ApplicationRestartState {
  if ($Mode -ne 'User') { return }
  $path = Get-ApplicationRestartStatePath
  if (Test-Path -LiteralPath $path) {
    $state = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    return $state.packages
  }
}

function Save-ApplicationRestartState {
  param([object[]]$Packages)
  if ($WhatIfPreference) { return }
  $path = Get-ApplicationRestartStatePath
  if ($Packages.Count -gt 0) {
    @{ schemaVersion = '1.0'; packages = @($Packages) } | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $path
  } elseif (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
  }
}

function Test-SameRecordedTarget {
  param([object]$Recorded, [object]$Package)
  if ($null -eq $Recorded -or $null -eq $Package) { return $false }
  return ($Recorded.id -eq $Package.Id -and $Recorded.source -eq $Package.Source -and $Recorded.available -eq $Package.Available)
}

function Add-PendingApplicationRestart {
  param([object]$Package, [string]$EvidencePath)
  $entries = @(Get-ApplicationRestartState | Where-Object { -not (Test-SameRecordedTarget $_ $Package) })
  $entries += [pscustomobject]@{
    id = [string]$Package.Id; name = [string]$Package.Name; source = [string]$Package.Source
    available = [string]$Package.Available; evidencePath = $EvidencePath
    runId = $script:RunId; expiresAt = (Get-Date).AddDays(7).ToString('o')
  }
  Save-ApplicationRestartState -Packages $entries
}

function Resolve-PendingApplicationRestart {
  param([object[]]$AvailablePackages)
  $pending = @(Get-ApplicationRestartState)
  $now = Get-Date
  $remaining = @(foreach ($entry in $pending) {
    $rule = Get-PackageRule -PackageId $entry.id
    $targetSatisfied = if (@(Get-PropertyValue $rule 'appxPackageFamilyNames' @()).Count) {
      # A pinned AppX target requires positive current-user Main registration evidence.
      $main = @(Get-MatchingUserAppxPackages -Package $entry)
      $main.Count -eq 1 -and (Compare-RecordedPackageVersion $entry.id ([string]$main[0].Version) $entry.available) -in @(0, 1)
    } else {
      Test-RecordedTargetInstalled -Package $entry -AvailablePackages $AvailablePackages
    }
    if ($targetSatisfied) {
      Write-AgentEvent -EventId 5100 -EntryType Information -Fields @{ packageId = $entry.id; status = 'ApplicationRestartVerificationResolved'; available = $entry.available; source = $entry.source; evidencePath = $entry.evidencePath } | Out-Null
    } elseif ([datetimeoffset]$entry.expiresAt -le $now) {
      Write-AgentEvent -EventId 5601 -EntryType Warning -Fields @{ packageId = $entry.id; status = 'ApplicationRestartVerificationExpired'; available = $entry.available; source = $entry.source; evidencePath = $entry.evidencePath; requiresAdminRemediation = $true } | Out-Null
    } else {
      $entry
    }
  })
  if ($remaining.Count -ne $pending.Count) { Save-ApplicationRestartState -Packages $remaining }
  return $remaining
}

function Write-ActiveRunState {
  if ($WhatIfPreference -or [string]::IsNullOrWhiteSpace($script:ActiveRunStatePath)) { return }

  [ordered]@{
    schemaVersion = '1.0'
    runId         = $script:RunId
    computerName  = $env:COMPUTERNAME
    mode          = $Mode
    processId     = $PID
    startedAt     = (Get-Date).ToString('o')
    deadline      = $script:RunDeadline.ToString('o')
    policyPath    = Resolve-PolicyPath
  } | ConvertTo-Json -Depth 4 | Write-AtomicJson -Path $script:ActiveRunStatePath
}

function Write-AtomicJson {
  param([Parameter(ValueFromPipeline = $true)][string]$Json, [string]$Path)
  process {
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
      [IO.File]::WriteAllText($temporary, $Json, (New-Object Text.UTF8Encoding($false)))
      if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temporary, $Path, [System.Management.Automation.Language.NullString]::Value) }
      else { [IO.File]::Move($temporary, $Path) }
    } finally { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
  }
}

function Enter-AgentRunLock {
  if ($WhatIfPreference) { return $true }
  try {
    $script:RunLock = [IO.File]::Open((Join-Path $LogRoot 'State\agent.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    return $true
  } catch [IO.IOException] { return $false }
}

function Clear-ActiveRunState {
  if ($WhatIfPreference -or [string]::IsNullOrWhiteSpace($script:ActiveRunStatePath)) { return }
  if (-not (Test-Path -LiteralPath $script:ActiveRunStatePath)) { return }
  $state = Get-Content -LiteralPath $script:ActiveRunStatePath -Raw | ConvertFrom-Json
  if ($state.runId -eq $script:RunId -and [int]$state.processId -eq $PID) {
    Remove-Item -LiteralPath $script:ActiveRunStatePath -Force -ErrorAction SilentlyContinue
  }
}

function Get-InapplicableUpgradeState {
  if ([string]::IsNullOrWhiteSpace($script:InapplicableUpgradeStatePath) -or -not (Test-Path -LiteralPath $script:InapplicableUpgradeStatePath)) { return @() }
  try {
    return @(Get-Content -LiteralPath $script:InapplicableUpgradeStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop | Select-Object -ExpandProperty packages)
  } catch {
    return @()
  }
}

function Save-InapplicableUpgradeState {
  param([object[]]$Packages)

  if ($WhatIfPreference) { return }
  [ordered]@{
    schemaVersion = '1.0'
    computerName  = $env:COMPUTERNAME
    mode          = $Mode
    packages      = @($Packages)
  } | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $script:InapplicableUpgradeStatePath
}

function Get-InapplicableUpgradeDeferral {
  param([object]$Package)

  $now = Get-Date
  return @(
    Get-InapplicableUpgradeState | Where-Object {
      $sameId = [string]::Equals([string]$_.id, [string]$Package.Id, [System.StringComparison]::OrdinalIgnoreCase)
      $sameInstalled = [string]::Equals([string]$_.installed, [string]$Package.Installed, [System.StringComparison]::OrdinalIgnoreCase)
      $sameAvailable = [string]::Equals([string]$_.available, [string]$Package.Available, [System.StringComparison]::OrdinalIgnoreCase)
      $sameSource = [string]::Equals([string]$_.source, [string]$Package.Source, [System.StringComparison]::OrdinalIgnoreCase)
      $sameScope = (Get-PropertyValue -Object $_ -Name 'scope' -Default (Get-WingetScopeArgument)) -eq (Get-WingetScopeArgument)
      $deferredUntil = $null
      try { $deferredUntil = [datetime]$_.deferredUntil } catch { }
      $reason = Get-PropertyValue $_ 'reason' ''
      $reason -notin @('KnownBadRelease','ManualOnly') -and $sameId -and $sameInstalled -and $sameAvailable -and $sameSource -and $sameScope -and $deferredUntil -and $deferredUntil -gt $now
    }
  ) | Select-Object -First 1
}

function Set-InapplicableUpgradeDeferral {
  param([object]$Package, [string]$EvidencePath, [object]$CommandOutcome)

  $existing = @(Get-InapplicableUpgradeState | Where-Object { -not [string]::Equals([string]$_.id, [string]$Package.Id, [System.StringComparison]::OrdinalIgnoreCase) })
  $deferredUntil = (Get-Date).AddDays([int]$script:Policy.inapplicableUpgradeDeferralDays)
  $existing += [pscustomobject]@{
    id            = [string]$Package.Id
    name          = [string]$Package.Name
    installed     = [string]$Package.Installed
    available     = [string]$Package.Available
    source        = [string]$Package.Source
    scope         = Get-WingetScopeArgument
    reason        = $CommandOutcome.Status
    deferredAt    = (Get-Date).ToString('o')
    deferredUntil = $deferredUntil.ToString('o')
    evidencePath  = $EvidencePath
    diagnostic    = [string]$CommandOutcome.Diagnostic
  }
  Save-InapplicableUpgradeState -Packages $existing
  return $deferredUntil
}

function Move-FailureNotificationToAdminDeferral {
  param([object]$Package, [string]$Reason, [string]$EvidencePath)

  if ($WhatIfPreference) { return }
  $path = Get-NotificationStatePath
  if (-not (Test-Path -LiteralPath $path)) { return }
  try { $state = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop } catch { return }
  $failed = @(Get-PropertyValue $state 'failedPackages' @())
  $matched = @($failed | Where-Object { $_.id -eq $Package.Id -and $_.source -eq $Package.Source -and $_.available -eq $Package.Available })
  if ($matched.Count -eq 0) { return }
  # Preserve the old failure under administrator-owned state before changing the user notice.
  $null = Set-InapplicableUpgradeDeferral -Package $Package -EvidencePath $EvidencePath -CommandOutcome ([pscustomobject]@{Status=$Reason;Diagnostic='Prior failure notification transferred to administrator review; not proof of installation.'})
  $deferrals = @(Get-InapplicableUpgradeState)
  foreach ($entry in $deferrals | Where-Object { $_.id -eq $Package.Id -and $_.source -eq $Package.Source -and $_.available -eq $Package.Available }) {
    $entry | Add-Member -NotePropertyName priorNotificationFailure -NotePropertyValue $matched -Force
  }
  Save-InapplicableUpgradeState -Packages $deferrals
  $state.failedPackages = @($failed | Where-Object { $_ -notin $matched })
  $state.summary.failed = $state.failedPackages.Count
  $restartCount = @(Get-PropertyValue $state 'restartRequiredPackages' @()).Count
  if ($state.failedPackages.Count -eq 0 -and $restartCount -eq 0) {
    Remove-Item -LiteralPath $path -Force
  } else {
    $state.status = if ($state.failedPackages.Count) { 'CompletedWithFailures' } else { 'CompletedRestartRequired' }
    $state | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $path
  }
  Write-AgentEvent -EventId 5101 -EntryType Warning -Fields @{packageId=$Package.Id;status='FailureReclassifiedForAdministrator';reason=$Reason;requiresAdminRemediation=$true;evidencePath=$EvidencePath} | Out-Null
}

function Clear-InapplicableUpgradeDeferral {
  param([object]$Package)

  $existing = @(Get-InapplicableUpgradeState)
  $remaining = @($existing | Where-Object { -not [string]::Equals([string]$_.id, [string]$Package.Id, [System.StringComparison]::OrdinalIgnoreCase) })
  if ($remaining.Count -ne $existing.Count) {
    Save-InapplicableUpgradeState -Packages $remaining
  }
}

function Get-SystemBootTime {
  try { return (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime } catch { return $null }
}

function Get-RebootMarkerSnapshot {
  $pendingRename = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
  $pendingRenameCount = if ($null -eq $pendingRename -or $null -eq $pendingRename.PSObject.Properties['PendingFileRenameOperations']) {
    0
  } else {
    @($pendingRename.PendingFileRenameOperations).Count
  }
  [pscustomobject]@{
    cbsRebootPending          = Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    windowsUpdateRebootPending = Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    pendingFileRenameCount    = $pendingRenameCount
    bootTime                  = Get-SystemBootTime
  }
}

function Test-NewRebootMarker {
  param([object]$Before, [object]$After)

  if ($null -eq $Before -or $null -eq $After) { return $false }
  return (
    ((-not [bool]$Before.cbsRebootPending) -and [bool]$After.cbsRebootPending) -or
    ((-not [bool]$Before.windowsUpdateRebootPending) -and [bool]$After.windowsUpdateRebootPending) -or
    ([int]$After.pendingFileRenameCount -gt [int]$Before.pendingFileRenameCount)
  )
}

function Get-PackageAttemptEvidencePath {
  param([object]$Package, [int]$Attempt = 1)
  $safeId = ([string]$Package.Id -replace '[^A-Za-z0-9._-]', '_')
  $safeRunId = $script:RunId -replace '[^A-Za-z0-9._-]', '_'
  return (Join-Path -Path (Join-Path -Path $LogRoot -ChildPath 'Logs') -ChildPath ("{0}-{1}-{2}-{3}.attempt.json" -f $safeId, $safeRunId, $Attempt, ([guid]::NewGuid().ToString('N'))))
}

function Write-PackageAttemptEvidence {
  param(
    [object]$Package,
    [object]$UpgradeResult,
    [Nullable[int]]$InstallerExitCode,
    [object]$InstallerOutcome,
    [object]$CommandOutcome,
    [object]$RebootBefore,
    [object]$RebootAfter,
    [string]$InstallerType,
    [string]$EvidencePath,
    [int]$Attempt,
    [string[]]$Arguments,
    [object]$Verification
  )

  if ($WhatIfPreference) { return }
  [ordered]@{
    schemaVersion      = '1.0'
    timestamp          = (Get-Date).ToString('o')
    computerName       = $env:COMPUTERNAME
    runId              = $script:RunId
    mode               = $Mode
    scope              = Get-WingetScopeArgument
    attempt            = $Attempt
    arguments          = @($Arguments)
    verification       = $Verification
    package            = [ordered]@{ id = [string]$Package.Id; name = [string]$Package.Name; installed = [string]$Package.Installed; available = [string]$Package.Available; source = [string]$Package.Source }
    installerType      = $InstallerType
    wingetExitCode     = [int]$UpgradeResult.ExitCode
    timedOut           = [bool]$UpgradeResult.TimedOut
    installerExitCode  = $InstallerExitCode
    installerOutcome   = $InstallerOutcome
    commandOutcome     = $CommandOutcome
    rebootBefore       = $RebootBefore
    rebootAfter        = $RebootAfter
    stdout             = [string]$UpgradeResult.StdOut
    stderr             = [string]$UpgradeResult.StdErr
  } | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $EvidencePath
}

function Add-PendingRestartVerification {
  param([object]$Package, [string]$EvidencePath, [string]$Reason)

  $statePath = Get-RestartVerificationStatePath
  $existing = @()
  if (Test-Path -LiteralPath $statePath) {
    try {
      $existingState = Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      $existing = @($existingState.packages)
    } catch { $existing = @() }
  }
  $remaining = @($existing | Where-Object { -not [string]::Equals([string]$_.id, [string]$Package.Id, [System.StringComparison]::OrdinalIgnoreCase) })
  $remaining += [pscustomobject]@{ id = [string]$Package.Id; name = [string]$Package.Name; source = [string]$Package.Source; installed = [string]$Package.Installed; available = [string]$Package.Available; detectedAt = (Get-Date).ToString('o'); bootTimeAtDetection = Get-SystemBootTime; reason = $Reason; evidencePath = $EvidencePath }
  [ordered]@{ schemaVersion = '1.0'; computerName = $env:COMPUTERNAME; mode = $Mode; packages = $remaining } | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $statePath
}

function Test-PackageTargetSatisfied {
  param([object]$Package, [object[]]$AvailablePackages)

  $offers = @($AvailablePackages | Where-Object {
    [string]::Equals([string]$_.Id, [string]$Package.Id, [StringComparison]::OrdinalIgnoreCase) -and
    [string]::Equals([string]$_.Source, [string]$Package.Source, [StringComparison]::OrdinalIgnoreCase)
  })
  if ($offers.Count -eq 0) { return $true }
  $target = $null
  if (-not [version]::TryParse([string]$Package.Available, [ref]$target)) { return $false }
  foreach ($offer in $offers) {
    $installed = $null
    if ([version]::TryParse([string]$offer.Installed, [ref]$installed) -and $installed -ge $target) { return $true }
  }
  return $false
}

function Resolve-PendingRestartVerification {
  param([object[]]$AvailablePackages)

  $statePath = Get-RestartVerificationStatePath
  $deferred = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $failed = New-Object System.Collections.Generic.List[object]
  $resolved = New-Object System.Collections.Generic.List[string]
  $review = New-Object System.Collections.Generic.List[object]
  $postBoot = New-Object System.Collections.Generic.List[object]
  if (-not (Test-Path -LiteralPath $statePath)) { return [pscustomobject]@{ DeferredPackageIds = @(); FailedPackages = @(); ResolvedPackageIds = @(); ReviewPackages = @(); PostBootPackages = @(); PendingRestartCount = 0 } }

  try {
    $pendingState = Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $pending = @($pendingState.packages)
  } catch {
    throw 'Restart verification state is unreadable; preserving it and stopping patching.'
  }
  $currentBootTime = Get-SystemBootTime
  $remaining = New-Object System.Collections.Generic.List[object]
  foreach ($entry in $pending) {
    $rebooted = $currentBootTime -and $entry.bootTimeAtDetection -and ([datetime]$currentBootTime -gt [datetime]$entry.bootTimeAtDetection)
    if (-not $rebooted) {
      [void]$deferred.Add([string]$entry.id)
      $remaining.Add($entry)
      Write-AgentEvent -EventId 5300 -EntryType Warning -Fields @{ packageId = [string]$entry.id; status = 'DeferredRestartVerificationPendingReboot'; restartRequired = $true; evidencePath = [string]$entry.evidencePath } | Out-Null
      continue
    }
    $postBoot.Add($entry)
    $dotNetProof = Get-DotNetRestartTargetProof -Package $entry
    $offers = @()
    if ($null -ne $dotNetProof) {
      # A WinGet row or an ARP MSI version can describe a different .NET architecture or component.
      $targetSatisfied = $dotNetProof.Verified
    } else {
      $targetSatisfied = Test-RecordedTargetInstalled -Package $entry -AvailablePackages $AvailablePackages
      $offers = @($AvailablePackages | Where-Object { $_.Id -eq $entry.id -and $_.Source -eq $entry.source })
    }
    if (-not $targetSatisfied -and $offers.Count -eq 1 -and (Get-PropertyValue $offers[0] 'InstallScope' '') -eq $Mode -and (Compare-RecordedPackageVersion $entry.id $offers[0].Installed $entry.available) -eq -1) {
      $failedPackage = $offers[0]
      $failed.Add($failedPackage)
      [void]$deferred.Add([string]$entry.id)
      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{ packageId = [string]$entry.id; status = 'RestartVerificationFailedStillOffered'; evidencePath = [string]$entry.evidencePath } | Out-Null
    } elseif ($targetSatisfied) {
      $resolved.Add([string]$entry.id)
      $fields = @{ packageId = [string]$entry.id; status = 'RestartVerificationResolved'; evidencePath = [string]$entry.evidencePath }
      if ($null -ne $dotNetProof) {
        $fields.verificationMethod = 'DotNetMachineRegistrationAndPayload'
        $fields.installed = [string]$entry.available
        $fields.architecture = $dotNetProof.Architecture
        $fields.verificationPath = $dotNetProof.Path
      }
      Write-AgentEvent -EventId 5100 -EntryType Information -Fields $fields | Out-Null
    } else {
      [void]$deferred.Add([string]$entry.id)
      $remaining.Add($entry)
      $review.Add($entry)
      Write-AgentEvent -EventId 5601 -EntryType Warning -Fields @{ packageId = [string]$entry.id; status = 'RestartVerificationInconclusive'; requiresAdminRemediation = $true; reason = 'InstalledTargetNotProvenAfterReboot'; evidencePath = [string]$entry.evidencePath } | Out-Null
    }
  }
  if ($remaining.Count -eq 0) {
    Remove-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue
  } else {
    [ordered]@{ schemaVersion = '1.0'; computerName = $env:COMPUTERNAME; mode = $Mode; packages = $remaining } | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $statePath
  }
  return [pscustomobject]@{
    DeferredPackageIds = @($deferred | ForEach-Object { [string]$_ })
    FailedPackages     = @($failed | ForEach-Object { $_ })
    ResolvedPackageIds  = @($resolved | ForEach-Object { $_ })
    ReviewPackages      = $review.ToArray()
    PostBootPackages    = $postBoot.ToArray()
    PendingRestartCount = $remaining.Count
  }
}

function Get-RegisteredDotNetInstallRoot {
  param([string]$Architecture, [string]$VersionKey, [string]$Version)
  $base = $null; $architectureKey = $null; $productKey = $null
  try {
    # The .NET host reads the machine-wide registration from the 32-bit registry view.
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry32')
    $architectureKey = $base.OpenSubKey("SOFTWARE\dotnet\Setup\InstalledVersions\$Architecture")
    if (-not $architectureKey) { return }
    $productKey = $architectureKey.OpenSubKey($VersionKey)
    if ($productKey -and $productKey.GetValue($Version, 0) -eq 1 -and $productKey.GetValueKind($Version) -eq 'DWord') {
      return $architectureKey.GetValue('InstallLocation')
    }
  } catch {
    return
  } finally {
    if ($productKey) { $productKey.Dispose() }
    if ($architectureKey) { $architectureKey.Dispose() }
    if ($base) { $base.Dispose() }
  }
}

function Get-DotNetRestartTargetProof {
  param([object]$Package)
  if ($Mode -ne 'Machine' -or $Package.source -ne 'winget' -or
    $Package.id -notmatch '^Microsoft\.DotNet\.(SDK|Runtime|AspNetCore|DesktopRuntime)\.([0-9]+)$') { return }
  $family = $Matches[1]
  $major = $Matches[2]
  $result = [pscustomobject]@{ Verified = $false; Architecture = $null; Path = $null }
  $version = [string]$Package.available
  if ($version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+$' -or ($version.Split('.')[0] -ne $major) -or
    $Package.name -notmatch '\((x64|x86|arm64)\)$') { return $result }
  $result.Architecture = $Matches[1].ToLowerInvariant()
  $framework, $binary = switch ($family) {
    'SDK'            { 'sdk'; 'dotnet.dll' }
    'Runtime'        { 'Microsoft.NETCore.App'; 'coreclr.dll' }
    'AspNetCore'     { 'Microsoft.AspNetCore.App'; 'Microsoft.AspNetCore.dll' }
    'DesktopRuntime' { 'Microsoft.WindowsDesktop.App'; 'PresentationFramework.dll' }
  }
  $registryPath = if ($family -eq 'SDK') { 'sdk' } else { "sharedfx\$framework" }
  $productPath = if ($family -eq 'SDK') { 'sdk' } else { "shared\$framework" }
  try {
    $root = Get-RegisteredDotNetInstallRoot -Architecture $result.Architecture -VersionKey $registryPath -Version $version
    if ($root -notmatch '^[a-z]:[\\/]') { return $result }
    $payload = Join-Path (Join-Path (Join-Path $root $productPath) $version) $binary
    if ((Test-Path -LiteralPath (Join-Path $root 'dotnet.exe') -PathType Leaf -ErrorAction Stop) -and
      (Test-Path -LiteralPath $payload -PathType Leaf -ErrorAction Stop)) {
      $result.Verified = $true
      $result.Path = $payload
    }
  } catch { return $result }
  return $result
}

function Compare-RecordedPackageVersion {
  param([string]$PackageId, [string]$Installed, [string]$Target)
  if ($PackageId -eq 'Zoom.Zoom.EXE' -and ($Installed -match '[()]' -or $Target -match '[()]')) {
    # Zoom exposes the same build as x.y.z (build) or x.y.z.build. Require both builds.
    if ($Installed -notmatch '^\d+\.\d+\.\d+(?:\.\d+| \(\d+\))$' -or $Target -notmatch '^\d+\.\d+\.\d+(?:\.\d+| \(\d+\))$') { return $null }
    $Installed = $Installed -replace ' \((\d+)\)$', '.$1'
    $Target = $Target -replace ' \((\d+)\)$', '.$1'
  }
  $installedVersion = $null
  $targetVersion = $null
  if ([version]::TryParse($Installed, [ref]$installedVersion) -and [version]::TryParse($Target, [ref]$targetVersion)) {
    return $installedVersion.CompareTo($targetVersion)
  }
  return $null
}

function Test-RecordedTargetInstalled {
  param([object]$Package, [object[]]$AvailablePackages)
  $matchedOffer = $false
  $offers = @($AvailablePackages | Where-Object { $_.Id -eq $Package.id -and $_.Source -eq $Package.source })
  if ($offers.Count -gt 1) { return $false }
  foreach ($offer in $offers) {
    $matchedOffer = $true
    if ((Get-PropertyValue $offer 'InstallScope' '') -eq $Mode -and (Compare-RecordedPackageVersion $Package.id $offer.Installed $Package.available) -in @(0, 1)) { return $true }
  }
  if ($matchedOffer) { return $false }
  # Absence from the upgrade list alone is not proof: require an exact, unique ARP name and scope.
  $name = Normalize-DisplayText -Text ([string]$Package.name)
  if (-not $name) { return $false }
  $entries = @(Get-CachedUninstallEntries | Where-Object { $_.DisplayNameNorm -eq $name })
  if ($entries.Count -gt 1) { return $false }
  if ($entries.Count -eq 1) {
    return ((Get-ScopeFromEntry $entries[0]) -eq $Mode -and (Compare-RecordedPackageVersion $Package.id $entries[0].DisplayVersion $Package.available) -in @(0, 1))
  }
  $appx = @(Get-MatchingUserAppxPackages -Package $Package)
  return ($appx.Count -eq 1 -and (Compare-RecordedPackageVersion $Package.id ([string]$appx[0].Version) $Package.available) -in @(0, 1))
}

function Resolve-RecordedPackageState {
  param([object[]]$AvailablePackages, [object[]]$PostBootRestartPackages, [object[]]$ApplicationRestartPackages)
  if ($WhatIfPreference) { return }
  $script:CachedUninstallEntries = $null
  $script:CachedUserAppxPackages = $null
  $statePath = Get-NotificationStatePath
  if (Test-Path -LiteralPath $statePath) {
    try {
      $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
      $previousFailed = @(Get-PropertyValue -Object $state -Name 'failedPackages' -Default @())
      $previousRestart = @(Get-PropertyValue -Object $state -Name 'restartRequiredPackages' -Default @())
      $failed = @(foreach ($entry in $previousFailed) {
        if (-not @($ApplicationRestartPackages | Where-Object { Test-SameRecordedTarget $entry $_ }).Count -and -not (Test-RecordedTargetInstalled -Package $entry -AvailablePackages $AvailablePackages)) { $entry }
      })
      $restart = @(foreach ($entry in $previousRestart) {
        # A later boot retires only the matching reboot notice; unresolved targets stay in admin state.
        if (-not @($PostBootRestartPackages | Where-Object { Test-SameRecordedTarget $entry $_ }).Count) { $entry }
      })
      if ($failed.Count -ne $previousFailed.Count -or $restart.Count -ne $previousRestart.Count) {
        $state.failedPackages = $failed
        $state.restartRequiredPackages = $restart
        $state.summary.failed = $failed.Count
        $state.summary.restartRequired = $restart.Count
        if ($failed.Count -eq 0 -and $restart.Count -eq 0) {
          Remove-Item -LiteralPath $statePath -Force
        } else {
          $state.status = if ($failed.Count) { 'CompletedWithFailures' } else { 'CompletedRestartRequired' }
          $state | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $statePath
        }
        Write-AgentEvent -EventId 5100 -EntryType Information -Fields @{ status = 'NotificationTargetsResolved'; failedRemaining = $failed.Count; restartRemaining = $restart.Count } | Out-Null
      }
    } catch { Write-Warning "Notification reconciliation skipped; preserving state: $($_.Exception.Message)" }
  }
  $deferrals = @(Get-InapplicableUpgradeState)
  $remaining = @($deferrals | Where-Object {
    $resolved = Test-RecordedTargetInstalled -Package $_ -AvailablePackages $AvailablePackages
    if ($resolved) { Write-AgentEvent -EventId 5100 -EntryType Information -Fields @{ packageId = [string]$_.id; status = 'DeferralResolvedExternally' } | Out-Null }
    -not $resolved
  })
  if ($remaining.Count -ne $deferrals.Count) { Save-InapplicableUpgradeState -Packages $remaining }
}

function Write-NotificationState {
  param(
    [string]$Status,
    [hashtable]$Summary,
    [object[]]$RestartRequiredPackages,
    [object[]]$FailedPackages
  )

  if ($Status -notin @('CompletedRestartRequired', 'CompletedWithFailures', 'PreflightFailed', 'MachineModeNotRunningAsSystem', 'UnhandledAgentError')) {
    return
  }

  $statePath = Get-NotificationStatePath
  if (Test-Path -LiteralPath $statePath) {
    # A failure in another package must not discard a still-pending reboot or unresolved target.
    $previous = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    $RestartRequiredPackages = @($RestartRequiredPackages) + @(Get-PropertyValue -Object $previous -Name 'restartRequiredPackages' -Default @() | Where-Object { $_.id -notin @($RestartRequiredPackages | ForEach-Object { $_.Id }) })
    $FailedPackages = @($FailedPackages) + @(Get-PropertyValue -Object $previous -Name 'failedPackages' -Default @() | Where-Object { $_.id -notin @($FailedPackages | ForEach-Object { $_.Id }) })
  }
  $succeededCount = 0
  $failedCount = 0
  if ($Summary.ContainsKey('Succeeded')) { $succeededCount = [int]$Summary.Succeeded }
  if ($Summary.ContainsKey('Failed')) { $failedCount = [int]$Summary.Failed }

  $state = [ordered]@{
    schemaVersion           = '1.0'
    source                  = 'DanmadePatchAgent'
    computerName            = $env:COMPUTERNAME
    mode                    = $Mode
    runId                   = $script:RunId
    status                  = $Status
    timestamp               = (Get-Date).ToString('o')
    summary                 = [ordered]@{
      succeeded       = $succeededCount
      restartRequired = @($RestartRequiredPackages).Count
      failed          = [Math]::Max($failedCount, @($FailedPackages).Count)
    }
    restartRequiredPackages = @(
      foreach ($pkg in @($RestartRequiredPackages)) {
        [ordered]@{
          id        = [string]$pkg.Id
          name      = [string]$pkg.Name
          available = [string]$pkg.Available
          source    = [string]$pkg.Source
        }
      }
    )
    failedPackages          = @(
      foreach ($pkg in @($FailedPackages)) {
        [ordered]@{
          id        = [string]$pkg.Id
          name      = [string]$pkg.Name
          available = [string]$pkg.Available
          source    = [string]$pkg.Source
        }
      }
    )
  }

  if ($PSCmdlet.ShouldProcess($statePath, 'Write pending notification state')) {
    $state | ConvertTo-Json -Depth 8 | Write-AtomicJson -Path $statePath
  }
}

function Clear-NotificationState {
  param([switch]$PreserveRestartRequired)

  $statePath = Get-NotificationStatePath
  if (-not (Test-Path -LiteralPath $statePath)) { return }
  if ($PreserveRestartRequired) {
    try {
      $state = Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      if (@(Get-PropertyValue -Object $state -Name 'restartRequiredPackages' -Default @()).Count -gt 0 -or @(Get-PropertyValue -Object $state -Name 'failedPackages' -Default @()).Count -gt 0 -or [string]$state.status -eq 'CompletedRestartRequired') {
        return
      }
    } catch {
      Write-Warning 'Notification state is unreadable; preserving it for administrator review.'
      return
    }
  }
  if ($PSCmdlet.ShouldProcess($statePath, 'Clear pending notification state')) {
    Remove-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue
  }
}

function Initialize-EventSource {
  if (-not $script:Policy.wazuhReporting.eventLog) { return }
  if (-not $PSCmdlet.ShouldProcess($script:EventSource, 'Ensure Windows Event Log source')) { return }

  try {
    if (-not [System.Diagnostics.EventLog]::SourceExists($script:EventSource)) {
      New-EventLog -LogName $script:EventLogName -Source $script:EventSource
    }
  } catch {
    $script:EventLogUnavailable = $_.Exception.Message
  }
}

function ConvertTo-EventJson {
  param([hashtable]$Fields)

  $base = [ordered]@{
    schemaVersion     = $script:SchemaVersion
    runId             = $script:RunId
    computerName      = $env:COMPUTERNAME
    mode              = $Mode
    packageId         = $null
    status            = $null
    wingetExitCode    = $null
    wingetExitCodeHex = $null
    installerExitCode = $null
    retryCount        = 0
    recoveryActions   = @()
    restartRequired   = $false
    logPath           = $null
    timestamp         = (Get-Date).ToString('o')
  }

  foreach ($key in $Fields.Keys) {
    $base[$key] = $Fields[$key]
  }

  return ($base | ConvertTo-Json -Compress -Depth 8)
}

function Write-AgentEvent {
  param(
    [int]$EventId,
    [ValidateSet('Information', 'Warning', 'Error')]
    [string]$EntryType = 'Information',
    [hashtable]$Fields
  )

  $payload = ConvertTo-EventJson -Fields $Fields

  if ($script:Policy.wazuhReporting.jsonl -and $script:JsonlPath -and $PSCmdlet.ShouldProcess($script:JsonlPath, "Append event $EventId")) {
    try {
      Add-Content -LiteralPath $script:JsonlPath -Value $payload -Encoding utf8
    } catch {
      Write-Warning "Unable to write JSONL patch-agent event: $($_.Exception.Message)"
    }
  }

  if ($script:Policy.wazuhReporting.eventLog -and -not $script:EventLogUnavailable -and $PSCmdlet.ShouldProcess($script:EventSource, "Write event $EventId")) {
    try {
      Write-EventLog -LogName $script:EventLogName -Source $script:EventSource -EventId $EventId -EntryType $EntryType -Message $payload
    } catch {
      $script:EventLogUnavailable = $_.Exception.Message
      Write-Warning "Unable to write Windows Event Log patch-agent event: $($_.Exception.Message)"
    }
  }

  Write-Output $payload
}

function Remove-OldLogs {
  if ($WhatIfPreference) { return }
  $days = [int]$script:Policy.logRetentionDays
  $cutoff = (Get-Date).AddDays(-1 * $days)
  foreach ($path in @((Join-Path $LogRoot 'Logs'), (Join-Path $LogRoot 'Events'), (Join-Path $LogRoot 'Inventory'))) {
    if (-not (Test-Path -LiteralPath $path)) { continue }
    Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue |
      Where-Object { $_.LastWriteTime -lt $cutoff } |
      Remove-Item -Force -ErrorAction SilentlyContinue
  }
}

function Test-MaintenanceWindow {
  $window = $script:Policy.maintenanceWindow
  if (-not $window.enabled) { return $true }
  if ($Mode -eq 'User' -and (Get-PropertyValue -Object $window -Name 'machineOnly' -Default $false)) { return $true }

  $start = [TimeSpan]::Zero
  $end = [TimeSpan]::Zero
  if (-not [TimeSpan]::TryParse([string]$window.start, [ref]$start)) { throw 'Invalid maintenance window start.' }
  if (-not [TimeSpan]::TryParse([string]$window.end, [ref]$end)) { throw 'Invalid maintenance window end.' }

  $now = (Get-Date).TimeOfDay
  if ($start -le $end) {
    return ($now -ge $start -and $now -le $end)
  }

  return ($now -ge $start -or $now -le $end)
}

function Test-IsSystemAccount {
  $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  return ($identity.User.Value -eq 'S-1-5-18')
}

function Resolve-WingetCommand {
  $winget = Get-Command winget -ErrorAction SilentlyContinue
  if ($winget -and (Test-Path -LiteralPath $winget.Path)) {
    return $winget.Path
  }

  try {
    $appInstaller = Get-AppxPackage -AllUsers -Name Microsoft.DesktopAppInstaller -ErrorAction SilentlyContinue |
      Sort-Object Version -Descending |
      Select-Object -First 1
    if ($appInstaller -and $appInstaller.InstallLocation) {
      $candidate = Join-Path -Path $appInstaller.InstallLocation -ChildPath 'winget.exe'
      if (Test-Path -LiteralPath $candidate) {
        return $candidate
      }
    }
  } catch {
    Write-Verbose "Unable to resolve winget from App Installer package metadata. $($_.Exception.Message)"
  }

  $windowsAppsRoot = Join-Path -Path $env:ProgramFiles -ChildPath 'WindowsApps'
  if (Test-Path -LiteralPath $windowsAppsRoot) {
    $candidate = Get-ChildItem -LiteralPath $windowsAppsRoot -Directory -Filter 'Microsoft.DesktopAppInstaller_*__8wekyb3d8bbwe' -ErrorAction SilentlyContinue |
      Sort-Object Name -Descending |
      ForEach-Object { Join-Path -Path $_.FullName -ChildPath 'winget.exe' } |
      Where-Object { Test-Path -LiteralPath $_ } |
      Select-Object -First 1
    if ($candidate) { return $candidate }
  }

  return $null
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
          DisplayName     = Get-OptionalPropertyValue -Object $_ -Name 'DisplayName'
          DisplayNameNorm = Normalize-DisplayText -Text (Get-OptionalPropertyValue -Object $_ -Name 'DisplayName')
          DisplayVersion  = Get-OptionalPropertyValue -Object $_ -Name 'DisplayVersion'
          InstallLocation = Get-OptionalPropertyValue -Object $_ -Name 'InstallLocation'
          UninstallString = Get-OptionalPropertyValue -Object $_ -Name 'UninstallString'
          PSPath          = [string]$_.PSPath
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
  if ($installLocation -and $installLocation -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($uninstallString -and $uninstallString -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($installLocation -and $installLocation -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $installLocation -and $installLocation -match "^$programFilesX86") { return 'Machine' }
  if ($uninstallString -and $uninstallString -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $uninstallString -and $uninstallString -match "^$programFilesX86") { return 'Machine' }

  return 'Unknown'
}

function Get-CachedUserAppxPackages {
  if ($Mode -ne 'User' -or (Test-IsSystemAccount)) { return @() }
  if ($null -ne $script:CachedUserAppxPackages) { return $script:CachedUserAppxPackages }
  try {
    # Query this user's main packages: bundle versions are not installed application versions.
    $script:CachedUserAppxPackages = @(Get-AppxPackage -PackageTypeFilter Main -ErrorAction Stop |
      Where-Object { [int]$_.Status -eq 0 -and -not $_.IsFramework -and -not $_.IsResourcePackage -and (Get-PropertyValue $_ 'IsBundle' $true) -eq $false })
  } catch {
    $script:CachedUserAppxPackages = @()
  }
  return $script:CachedUserAppxPackages
}

function Get-MatchingUserAppxPackages {
  param([object]$Package)
  $rule = Get-PackageRule -PackageId ([string]$Package.Id)
  if (-not $rule) { return @() }
  $families = @(Get-PropertyValue $rule 'appxPackageFamilyNames' @())
  if ($families.Count -eq 0) { return @() }
  return @(Get-CachedUserAppxPackages | Where-Object { $families -contains [string]$_.PackageFamilyName })
}

function Resolve-PackageInstallScope {
  param([object]$Package)

  if (-not $Package -or [string]::IsNullOrWhiteSpace([string]$Package.Name)) {
    return 'Unknown'
  }

  $packageNameNorm = Normalize-DisplayText -Text ([string]$Package.Name)
  if (-not $packageNameNorm) { return 'Unknown' }

  $bestMatch = $null
  $bestMatchCount = 0
  $bestScore = -1
  foreach ($entry in @(Get-CachedUninstallEntries)) {
    if (-not $entry.DisplayNameNorm) { continue }

    $score = -1
    if ($entry.DisplayName -eq [string]$Package.Name) {
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
      $bestMatchCount = 1
    } elseif ($score -eq $bestScore) {
      $bestMatchCount++
    }
  }

  if ($bestMatchCount -gt 1) { return 'Unknown' }
  if (-not $bestMatch) {
    if (@(Get-MatchingUserAppxPackages -Package $Package).Count -eq 1) { return 'User' }
    return 'Unknown'
  }
  if ((Get-ScopeFromEntry -Entry $bestMatch) -ne 'User' -and @(Get-MatchingUserAppxPackages -Package $Package).Count -gt 0) { return 'Unknown' }
  return Get-ScopeFromEntry -Entry $bestMatch
}

function Add-PackageScope {
  param([object]$Package)

  $scope = Resolve-PackageInstallScope -Package $Package
  $Package | Add-Member -NotePropertyName InstallScope -NotePropertyValue $scope -Force
  return $Package
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

function Test-IsPlainVersionToken {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
  return $Text.Trim() -match '^\d+(?:\.\d+)+(?:[-+._][A-Za-z0-9]+)?$'
}

function ConvertFrom-WingetUpgradeTableRow {
  param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

  $idMatches = [regex]::Matches($Line, '(?<=^|\s)[A-Za-z0-9][A-Za-z0-9._+-]*(?=\s)')
  for ($i = $idMatches.Count - 1; $i -ge 0; $i--) {
    $candidate = $idMatches[$i]
    if (Test-IsPlainVersionToken -Text $candidate.Value) { continue }

    $name = $Line.Substring(0, $candidate.Index).Trim()
    $tail = $Line.Substring($candidate.Index + $candidate.Length).Trim()
    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($tail)) { continue }

    $tailMatch = [regex]::Match(
      $tail,
      '^(?<Installed>(?:[<>]=?\s*)?\S+(?:\s+\([^)]+\))?)\s+(?<Available>(?:[<>]=?\s*)?\S+(?:\s+\([^)]+\))?)\s+(?<Source>\S+)\s*$'
    )
    if (-not $tailMatch.Success) { continue }

    $id = Normalize-WingetPackageId -RawId $candidate.Value
    if (-not $id) { continue }

    return [pscustomobject]@{
      Name      = $name
      Id        = $id
      Installed = $tailMatch.Groups['Installed'].Value.Trim()
      Available = $tailMatch.Groups['Available'].Value.Trim()
      Source    = $tailMatch.Groups['Source'].Value.Trim()
    }
  }

  return $null
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
    if (-not [string]::IsNullOrWhiteSpace($text)) { return $text.Trim() }
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
    $candidateSource = Get-FirstValue -Object $Node -Names @('Source')
    if ($candidateSource) { $sourceFromNode = $candidateSource }
  } elseif (($nodeProps -contains 'Name') -and ($nodeProps -contains 'Packages')) {
    $candidateSource = Get-FirstValue -Object $Node -Names @('Name')
    if ($candidateSource) { $sourceFromNode = $candidateSource }
  } elseif ($nodeProps -contains 'Details') {
    $candidateSource = Get-FirstValue -Object $Node.Details -Names @('Name')
    if ($candidateSource) { $sourceFromNode = $candidateSource }
  }

  $id = Normalize-WingetPackageId -RawId (Get-FirstValue -Object $Node -Names @('PackageIdentifier', 'Id'))
  if ($id) {
    $name = Get-FirstValue -Object $Node -Names @('PackageName', 'Name')
    if (-not $name) { $name = $id }

    $installed = Get-FirstValue -Object $Node -Names @('InstalledVersion', 'Version', 'Installed')
    $available = Get-FirstValue -Object $Node -Names @('AvailableVersion', 'Available', 'LatestVersion')
    $source = Get-FirstValue -Object $Node -Names @('Source', 'Repository', 'Origin')
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

function Get-WingetInventory {
  $baseArgs = @('upgrade', '--disable-interactivity', '--accept-source-agreements', '--scope', (Get-WingetScopeArgument))
  if ($script:Policy.includeUnknown) { $baseArgs += '--include-unknown' }
  $evidencePath = Join-Path $LogRoot ("Inventory\query-{0}.json" -f [guid]::NewGuid().ToString('N'))
  $queries = New-Object System.Collections.Generic.List[object]
  $packages = @()
  $status = 'Failed'
  $message = $null
  try {
    $jsonResult = Invoke-WingetProcess -Arguments ($baseArgs + @('--output', 'json')) -TimeoutSeconds 180
    $queries.Add($jsonResult)
    $output = ([string]$jsonResult.StdOut, [string]$jsonResult.StdErr) -join [Environment]::NewLine
    if ($jsonResult.TimedOut) { throw 'JSON inventory command timed out.' }
    # Only a positively identified unsupported option permits text fallback.
    $unsupported = $output -match '(?i)(unrecognized|unknown|not recognized|invalid).*(argument|option).*output|argument.*not recognized.*output'
    if (-not $unsupported) {
      if ($jsonResult.TimedOut -or [int]$jsonResult.ExitCode -ne 0) { throw 'JSON inventory command failed or timed out.' }
      $jsonText = ([string]$jsonResult.StdOut).Trim()
      $first = $jsonText.IndexOfAny(@('[', '{'))
      if ($first -lt 0) { throw 'JSON inventory did not contain a recognized document.' }
      $jsonText = $jsonText.Substring($first)
      $data = ConvertFrom-Json -InputObject $jsonText -ErrorAction Stop
      $collected = New-Object System.Collections.Generic.List[object]
      Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null
      if ($collected.Count -eq 0 -and $jsonText -notmatch '^\[\s*\]$') {
        throw 'JSON inventory schema was unrecognized or contained no usable package records.'
      }
      if (@($collected | Where-Object { -not $_.Installed -or -not $_.Available }).Count -gt 0) { throw 'JSON inventory contained incomplete versions.' }
      $packages = @($collected.ToArray())
    } else {
      $tableResult = Invoke-WingetProcess -Arguments $baseArgs -TimeoutSeconds 180
      $queries.Add($tableResult)
      $tableOutput = ([string]$tableResult.StdOut, [string]$tableResult.StdErr) -join [Environment]::NewLine
      if ($tableOutput -match '(?i)Failed when searching source|Failed to open.*source|An unexpected error occurred') { throw 'Text inventory reported a source or command failure.' }
      $empty = [string]$tableResult.StdOut -match '(?im)^No available upgrade found\.?\s*$|^No installed package found matching input criteria\.?\s*$'
      $emptyCode = ([string]$tableResult.StdOut -match '(?im)^No available upgrade found' -and (Convert-ToHexCode $tableResult.ExitCode) -eq '0x8A15002B') -or
        ([string]$tableResult.StdOut -match '(?im)^No installed package found matching input criteria' -and (Convert-ToHexCode $tableResult.ExitCode) -eq '0x8A150014')
      if ($tableResult.TimedOut -or ([int]$tableResult.ExitCode -ne 0 -and -not ($empty -and $emptyCode))) { throw 'Text inventory command failed or timed out.' }
      if (-not $empty) {
        $lines = [string]$tableResult.StdOut -split "\r\n|\n|\r"
        $header = $lines | Where-Object { $_ -match '^Name\s+Id\s+Version\s+Available\s+Source' } | Select-Object -First 1
        if (-not $header) { throw 'Text inventory header was not recognized.' }
        $headerIndex = [array]::IndexOf($lines, $header)
        $packages = @(
          foreach ($line in ($lines | Select-Object -Skip ($headerIndex + 2))) {
            if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^[-\s]+$|^\d+\s+upgrades available\.?$|^\d+ package\(s\).*|^The following packages') { continue }
            $pkg = ConvertFrom-WingetUpgradeTableRow -Line $line
            if (-not $pkg) { throw "Unrecognized inventory row: $line" }
            $pkg
          }
        )
        if ($packages.Count -eq 0) { throw 'An empty table is not a confirmed no-upgrades result.' }
      }
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $packages = @(foreach ($pkg in $packages) { if ($seen.Add("$($pkg.Id)|$($pkg.Source)")) { Add-PackageScope -Package $pkg } })
    $status = if ($packages.Count -eq 0) { 'Empty' } else { 'Succeeded' }
  } catch { $message = $_.Exception.Message }
  $result = [pscustomobject]@{
    Status = $status
    Packages = @($packages)
    Message = $message
    EvidencePath = $evidencePath
    Scope = Get-WingetScopeArgument
    Queries = @($queries.ToArray())
  }
  if (-not $WhatIfPreference) { $result | ConvertTo-Json -Depth 10 | Write-AtomicJson -Path $evidencePath }
  return $result
}

function Get-WingetUpgradeList {
  param([switch]$FailOnError)
  # All callers must fail closed; retain the switch for older diagnostics.
  $inventory = Get-WingetInventory
  if ($inventory.Status -eq 'Failed') {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
      status = 'InventoryInconclusive'
      message = $inventory.Message
      evidencePath = $inventory.EvidencePath
    } | Out-Null
    throw "Inventory inconclusive: $($inventory.Message)"
  }
  return @($inventory.Packages)
}

function Get-PackageRule {
  param([string]$PackageId)

  foreach ($rule in @($script:Policy.packageRules)) {
    if ([string]::Equals([string]$rule.id, $PackageId, [System.StringComparison]::OrdinalIgnoreCase)) {
      return $rule
    }
  }
  return $null
}

function Get-PackageMajorVersion {
  param([string]$Version)

  $match = [regex]::Match([string]$Version, '(?<!\d)(\d+)')
  if (-not $match.Success) { return $null }
  return [int]$match.Groups[1].Value
}

function Test-PackageMajorUpgrade {
  param([object]$Package)

  $installedMajor = Get-PackageMajorVersion -Version ([string]$Package.Installed)
  $availableMajor = Get-PackageMajorVersion -Version ([string]$Package.Available)
  return ($null -ne $installedMajor -and $null -ne $availableMajor -and $installedMajor -ne $availableMajor)
}

function Get-RunningRuleProcesses {
  param([object]$Rule)

  if ($null -eq $Rule -or @($Rule.processNames).Count -eq 0) { return @() }
  $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($name in @($Rule.processNames)) {
    $trimmed = ([string]$name).Trim()
    if ($trimmed.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) { $trimmed = $trimmed.Substring(0, $trimmed.Length - 4) }
    if ($trimmed) { [void]$names.Add($trimmed) }
  }
  return @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $names.Contains($_.ProcessName) })
}

function Test-RetryableInstallerExitCode {
  param([Nullable[int]]$InstallerExitCode)

  if ($null -eq $InstallerExitCode) { return $false }
  return @($script:Policy.retryableInstallerExitCodes) -contains [int]$InstallerExitCode
}

function Select-PolicyPackages {
  param(
    [object[]]$Packages,
    [object[]]$DeferredPackageIds,
    [object[]]$ApplicationRestartPackages
  )

  $script:SelectionDeferredCount = 0
  $script:SelectionCounts = @{ blockedCount = 0; notAllowedCount = 0; nativeUpdaterCount = 0; manualOnlyCount = 0; majorVersionDeferredCount = 0; processInUseCount = 0; scopeUnknownCount = 0; scopeMismatchCount = 0; restartDeferredCount = 0 }
  $deferred = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($id in @($DeferredPackageIds)) { if ($id) { [void]$deferred.Add([string]$id) } }

  $allowed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($id in @($script:Policy.allowedPackageIds)) { [void]$allowed.Add($id) }

  $blocked = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($id in @($script:Policy.blockedPackageIds)) { [void]$blocked.Add($id) }

  return @(
    foreach ($pkg in @($Packages)) {
      if (-not $pkg.Id) { continue }
      if ($deferred.Contains([string]$pkg.Id)) { $script:SelectionDeferredCount++; $script:SelectionCounts.restartDeferredCount++; continue }
      if ($blocked.Contains([string]$pkg.Id)) {
        $script:SelectionCounts.blockedCount++
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedBlocked'
        } | Out-Null
        continue
      }
      if ($allowed.Count -gt 0 -and -not $allowed.Contains([string]$pkg.Id)) {
        $script:SelectionCounts.notAllowedCount++
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedNotAllowed'
        } | Out-Null
        continue
      }
      $rule = Get-PackageRule -PackageId ([string]$pkg.Id)
      if ($rule -and $rule.mode -eq 'NativeUpdater') {
        $script:SelectionCounts.nativeUpdaterCount++
        $script:SelectionDeferredCount++
        Write-AgentEvent -EventId 5101 -EntryType Warning -Fields @{
          packageId = [string]$pkg.Id; status = 'DeferredNativeUpdater'
          installed = [string]$pkg.Installed; available = [string]$pkg.Available
          source = [string]$pkg.Source; requiresAdminRemediation = $true
        } | Out-Null
        continue
      }
      if ($rule -and $rule.mode -eq 'ManualOnly') {
        $script:SelectionCounts.manualOnlyCount++
        $script:SelectionDeferredCount++
        Move-FailureNotificationToAdminDeferral -Package $pkg -Reason 'ManualOnly'
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'DeferredManualOnly'
          requiresAdminRemediation = $true
        } | Out-Null
        continue
      }

      $allowMajorVersion = if ($rule) { [bool]$rule.allowMajorVersion } else { $false }
      if ((Test-PackageMajorUpgrade -Package $pkg) -and -not $allowMajorVersion) {
        $script:SelectionCounts.majorVersionDeferredCount++
        $script:SelectionDeferredCount++
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'DeferredMajorVersion'
          installedVersion = [string]$pkg.Installed
          availableVersion = [string]$pkg.Available
        } | Out-Null
        continue
      }

      if ($rule -and $rule.mode -eq 'QuiescenceRequired') {
        $runningProcesses = @(Get-RunningRuleProcesses -Rule $rule)
        if ($runningProcesses.Count -gt 0) {
          $script:SelectionCounts.processInUseCount++
          $script:SelectionDeferredCount++
          Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
            packageId = [string]$pkg.Id
            status    = 'DeferredProcessInUse'
            processCount = $runningProcesses.Count
            processNames = @($runningProcesses | Select-Object -ExpandProperty ProcessName -Unique)
          } | Out-Null
          continue
        }
      }

      $scope = if ($pkg.PSObject.Properties['InstallScope']) { [string]$pkg.InstallScope } else { 'Unknown' }
      if ($Mode -eq 'User' -and $scope -ne 'User') {
        if ($scope -eq 'Unknown') { $script:SelectionCounts.scopeUnknownCount++ } else { $script:SelectionCounts.scopeMismatchCount++ }
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'InventoryScopeMismatch'
          installScope = $scope
          expectedScope = 'User'
        } | Out-Null
        continue
      }
      if ($Mode -eq 'Machine' -and $scope -eq 'User') {
        $script:SelectionCounts.scopeMismatchCount++
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'InventoryScopeMismatch'
          installScope = $scope
          expectedScope = 'Machine'
        } | Out-Null
        continue
      }
      if (@($ApplicationRestartPackages | Where-Object { Test-SameRecordedTarget $_ $pkg }).Count) {
        $script:SelectionDeferredCount++
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{ packageId = $pkg.Id; status = 'DeferredApplicationRestartPendingVerification'; available = $pkg.Available; source = $pkg.Source } | Out-Null
        continue
      }

      $pkg
    }
  )
}

function Get-AssessmentSummaryFields {
  param([int]$DiscoveredCount, [int]$SelectedCount, [hashtable]$Summary, [int]$CarriedRestartCount, [int]$PostRebootReviewCount, [int]$ApplicationRestartPendingCount)
  $pending = $CarriedRestartCount + $Summary.RestartRequired
  $fields = @{
    assessmentCompleted = $true
    assessmentOutcome = $(if ($DiscoveredCount -eq 0) { 'NoUpdatesOffered' } elseif ($SelectedCount -eq 0) { 'NoEligibleUpdates' } else { 'UpdatesProcessed' })
    coverageStatus = $(if ($SelectedCount -lt $DiscoveredCount -or $Summary.Failed -gt 0 -or $Summary.Deferred -gt 0 -or $pending -gt 0 -or $ApplicationRestartPendingCount -gt 0) { 'UnresolvedOffers' } else { 'NoUnresolvedOffers' })
    discoveredCount = $DiscoveredCount; selectedCount = $SelectedCount; deferredCount = $Summary.Deferred
    succeededCount = $Summary.Succeeded; failedCount = $Summary.Failed
    applicabilityDeferredCount = $script:ApplicabilityDeferredCount
    restartRequired = ($pending -gt $PostRebootReviewCount); pendingRestartCount = $pending
    newRestartCount = $Summary.RestartRequired; carriedRestartCount = $CarriedRestartCount
    postRebootReviewCount = $PostRebootReviewCount; applicationRestartPendingCount = $ApplicationRestartPendingCount
  }
  foreach ($key in $script:SelectionCounts.Keys) { $fields[$key] = $script:SelectionCounts[$key] }
  return $fields
}

function ConvertFrom-InstallerLogText {
  param([string]$Text)
  $result = [ordered]@{ ExitCode = $null; FinalSuccess = $false; FailurePattern = $null; IgnoredSignals = @() }
  $customActionFailure = $null
  foreach ($line in ($Text -split '\r\n|\n|\r')) {
    # Burn's final bundle exit is authoritative, including HRESULT-wrapped MSI failures.
    if ($line -match '(?i)\bi007:\s*Exit code:\s*(0x[0-9a-f]+|-?\d+)\b') {
      $raw = $Matches[1]
      $code = if ($raw.StartsWith('0x', [StringComparison]::OrdinalIgnoreCase)) { [Convert]::ToInt32($raw.Substring(2), 16) } else { [int]$raw }
      if ((Convert-ToHexCode -Code $code) -match '^0x8007([0-9A-F]{4})$') { $code = [Convert]::ToInt32($Matches[1], 16) }
      $result.ExitCode = $code
      $result.FinalSuccess = $code -in @(0, 3010, 1641)
      $result.FailurePattern = if ($result.FinalSuccess) { $null } else { $line.Trim() }
      continue
    }
    if ($line -match '(?i)returned\s+actual\s+error\s+code\s+(-?\d+)') {
      if ($line -match '(?i)will be translated to success|continue marking|continue\s+on\s+error') {
        $result.IgnoredSignals += $line.Trim()
      } elseif ([int]$Matches[1] -ne 0) {
        $customActionFailure = [pscustomobject]@{ Code = [int]$Matches[1]; Line = $line.Trim() }
      }
      continue
    }
    # Terminal records are authoritative in chronological order, including nonzero status.
    if ($line -match '(?i)Installation success or error status:\s*(-?\d+)\b|MainEngineThread is returning\s+(-?\d+)\b|(?:install|installer|uninstall)\s+failed\s+with\s+exit\s+code:\s*(-?\d+)|installer\s+return\s+code\s*[:=]\s*(-?\d+)|msi(?:\s+installer)?\s+(?:return|exit)\s+code\s*[:=]\s*(-?\d+)|errorcode\s*:\s*(-?\d+)') {
      $code = @($Matches.GetEnumerator() | Where-Object { $_.Key -match '^[1-6]$' -and $_.Value -ne '' } | Select-Object -ExpandProperty Value)[0]
      $result.ExitCode = [int]$code
      $result.FinalSuccess = [int]$code -in @(0, 3010, 1641)
      $result.FailurePattern = if ($result.FinalSuccess) { $null } else { $line.Trim() }
    } elseif ($line -match '(?i)Installation completed successfully\.?') {
      $result.ExitCode = 0
      $result.FinalSuccess = $true
      $result.FailurePattern = $null
    }
  }
  if ($null -eq $result.ExitCode -and $customActionFailure) {
    $result.ExitCode = $customActionFailure.Code
    $result.FailurePattern = $customActionFailure.Line
  }
  return [pscustomobject]$result
}

function Get-InstallerOutcomeFromLog {
  param([string]$Path)
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return ConvertFrom-InstallerLogText -Text '' }
  return ConvertFrom-InstallerLogText -Text (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop)
}

function Get-WingetCommandOutcome {
  param([object]$UpgradeResult)
  $output = ([string]$UpgradeResult.StdOut, [string]$UpgradeResult.StdErr) -join [Environment]::NewLine
  $urlMatch = [regex]::Match($output, '(?im)Downloading\s+(https?://\S+)')
  $url = if ($urlMatch.Success) { $urlMatch.Groups[1].Value.Trim() } else { $null }
  $hresult = [regex]::Match($output, '(?i)0x[0-9a-f]{8}').Value
  $httpStatus = $null
  $httpMatch = [regex]::Match($output, '(?i)0x8019([0-9a-f]{4})')
  if ($httpMatch.Success) { $httpStatus = [Convert]::ToInt32($httpMatch.Groups[1].Value, 16) }
  elseif ($output -match '(?i)HTTP(?:\s+status)?\s*[:=]?\s*(\d{3})\b') { $httpStatus = [int]$Matches[1] }

  $status = 'WingetFailed'
  $exitHex = Convert-ToHexCode -Code ([int]$UpgradeResult.ExitCode)
  $retryable = $false
  $installerStarted = $output -match '(?i)Starting package install|Starting installer|Successfully installed'
  $downloadFailure = $output -match '(?i)Download request status is not success|Package download failed|Failed when downloading|InternetReadFile\(\) failed|0x80072(ee2|efd|efe|ee7)|0x8019[0-9a-f]{4}'
  if ($UpgradeResult.TimedOut) { $status = 'InstallationOutcomeUncertain' }
  elseif ($downloadFailure) {
    $transient = $httpStatus -eq 429 -or ($httpStatus -ge 500 -and $httpStatus -lt 600) -or $output -match '(?i)0x80072(ee2|efd|efe|ee7)'
    $retryable = $transient -and -not $installerStarted
    $status = if ($retryable) { 'TransientDownloadFailure' } else { 'DownloadFailure' }
  } elseif (-not $installerStarted -and ($exitHex -eq '0x8A150010' -or ([int]$UpgradeResult.ExitCode -eq 0 -and $output -match '(?im)^No applicable installer found[;.]'))) {
    $status = 'NoApplicableInstaller'
  } elseif ($exitHex -eq '0x8A150014' -or ([int]$UpgradeResult.ExitCode -eq 0 -and $output -match '(?im)^No installed package found matching input criteria\.?\s*$')) {
    $status = 'InstalledPackageNotFound'
  } elseif ($exitHex -eq '0x8A15008E' -or ([int]$UpgradeResult.ExitCode -eq 0 -and $output -match '(?i)install technology is different|installer type.*different')) {
    $status = 'InstallerTechnologyMismatch'
  } elseif ($exitHex -in @('0x00000000','0x8A15002B') -and $output -match '(?i)No available upgrade found\.|No newer package versions are available') {
    $status = 'NoAvailableUpgrade'
  } elseif ($exitHex -eq '0x8A15002B' -or ([int]$UpgradeResult.ExitCode -eq 0 -and $output -match '(?i)No applicable upgrade found\.|does not apply to your system or requirements')) {
    $status = 'NoApplicableUpgrade'
  } elseif ($exitHex -in @('0x8A150109', '0x8A15010B') -or [int]$UpgradeResult.ExitCode -in @(3010,1641)) {
    $status = 'RestartRequired'
  } elseif ([int]$UpgradeResult.ExitCode -in @(0, 3010, 1641)) {
    $status = if ($output -match '(?i)Successfully (installed|upgraded)|Installation completed successfully') { 'Succeeded' } else { 'UnconfirmedOutcome' }
  }
  return [pscustomobject]@{
    Status = $status
    IsRetryable = $retryable
    HttpStatus = $httpStatus
    HResult = $hresult
    DownloadUrl = $url
    InstallerStarted = $installerStarted
    ApplicationRestartRequested = ($status -eq 'Succeeded' -and $output -match '(?im)\bRestart the application to complete the upgrade\.\s*$')
    Diagnostic = $output.Trim()
  }
}

function Wait-DownloadRetry {
  param([int]$Attempt)
  $delaySeconds = [math]::Min(60, 5 * [math]::Pow(2, $Attempt - 1)) + (Get-Random -Minimum 0 -Maximum 4)
  $remaining = Get-RemainingRunSeconds
  if ($null -ne $remaining -and $remaining -le $delaySeconds) { return $false }
  Start-Sleep -Seconds $delaySeconds
  return $true
}

function Test-RestartEvidence {
  param(
    [Nullable[int]]$InstallerExitCode,
    [object]$UpgradeResult,
    [object]$RebootBefore,
    [object]$RebootAfter
  )

  if ($InstallerExitCode -in @(3010, 1641)) { return "InstallerExitCode$InstallerExitCode" }
  $output = "{0}`n{1}" -f [string]$UpgradeResult.StdOut, [string]$UpgradeResult.StdErr
  if ((Convert-ToHexCode -Code ([int]$UpgradeResult.ExitCode)) -in @('0x8A150109','0x8A15010B')) { return 'WingetRestartExitCode' }
  if ([int]$UpgradeResult.ExitCode -in @(3010,1641)) { return 'WingetRestartExitCode' }
  if ($output -match '(?i)(restart|reboot)\s+(is\s+)?required|success.*reboot|Restart your PC to finish installation\.|Your PC will restart to finish installation\.') { return 'WingetRestartMessage' }
  if (Test-NewRebootMarker -Before $RebootBefore -After $RebootAfter) { return 'NewRebootMarker' }
  return $null
}

function Test-WingetPackageStillOffered {
  param([object]$Package)

  # `winget upgrade --id` executes an installer. Re-list the complete inventory instead.
  try {
    $availablePackages = @(Get-WingetUpgradeList -FailOnError)
  } catch {
    return [pscustomobject]@{
      Status            = 'VerificationInconclusive'
      StillOffered      = $false
      TimedOut          = $false
      WingetExitCode    = -1
      WingetExitCodeHex = Convert-ToHexCode -Code -1
    }
  }

  $stillOffered = -not (Test-PackageTargetSatisfied -Package $Package -AvailablePackages $availablePackages)

  if ($stillOffered) {
    return [pscustomobject]@{
      Status            = 'VerificationFailedStillOffered'
      StillOffered      = $true
      TimedOut          = $false
      WingetExitCode    = 0
      WingetExitCodeHex = Convert-ToHexCode -Code 0
    }
  }

  return [pscustomobject]@{
    Status            = 'VerifiedNotOffered'
    StillOffered      = $false
    TimedOut          = $false
    WingetExitCode    = 0
    WingetExitCodeHex = Convert-ToHexCode -Code 0
  }
}

function Invoke-WingetSourceUpdate {
  param([switch]$AfterReset)

  if (-not $PSCmdlet.ShouldProcess('winget source update', 'Run winget source update')) {
    return $true
  }

  $result = Invoke-WingetProcess -Arguments @('source', 'update') -TimeoutSeconds 180
  $exitCode = [int]$result.ExitCode
  if ($result.TimedOut) {
    Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
      status            = 'WingetSourceUpdateTimedOut'
      wingetExitCode    = $exitCode
      wingetExitCodeHex = Convert-ToHexCode -Code $exitCode
      recoveryActions   = @('sourceUpdate')
    } | Out-Null
    return $false
  }
  if ($exitCode -eq 0) {
    if ($AfterReset) {
      Write-AgentEvent -EventId 5500 -EntryType Information -Fields @{
        status          = 'WingetSourceRepairSucceeded'
        recoveryActions = @('sourceReset', 'sourceUpdate')
      } | Out-Null
    }
    return $true
  }

  Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
    status          = 'WingetSourceUpdateFailed'
    wingetExitCode  = $exitCode
    wingetExitCodeHex = Convert-ToHexCode -Code $exitCode
    recoveryActions = @('sourceUpdate')
  } | Out-Null
  return $false
}

function Invoke-WingetSourceReset {
  if (-not $script:Policy.wingetSourceRepair) { return $false }
  if (-not $PSCmdlet.ShouldProcess('winget source reset --force', 'Run winget source reset')) {
    return $true
  }

  $result = Invoke-WingetProcess -Arguments @('source', 'reset', '--force') -TimeoutSeconds 180
  $exitCode = [int]$result.ExitCode
  if ($result.TimedOut) {
    Write-AgentEvent -EventId 5500 -EntryType Error -Fields @{
      status            = 'WingetSourceRepairTimedOut'
      wingetExitCode    = $exitCode
      wingetExitCodeHex = Convert-ToHexCode -Code $exitCode
      recoveryActions   = @('sourceReset')
    } | Out-Null
    return $false
  }
  if ($exitCode -eq 0) {
    Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
      status          = 'WingetSourceResetAttempted'
      recoveryActions = @('sourceReset')
    } | Out-Null
    return Invoke-WingetSourceUpdate -AfterReset
  }

  Write-AgentEvent -EventId 5500 -EntryType Error -Fields @{
    status            = 'WingetSourceRepairFailed'
    wingetExitCode    = $exitCode
    wingetExitCodeHex = Convert-ToHexCode -Code $exitCode
    recoveryActions   = @('sourceReset')
  } | Out-Null
  return $false
}

function Test-WingetPreflight {
  $wingetPath = Resolve-WingetCommand
  if ([string]::IsNullOrWhiteSpace($wingetPath)) {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
      status = 'WingetNotFound'
    } | Out-Null
    return $false
  }

  $script:WingetCommand = $wingetPath

  if ($PSCmdlet.ShouldProcess('winget --info', 'Verify winget health')) {
    $infoResult = Invoke-WingetProcess -Arguments @('--info') -TimeoutSeconds 60
    $infoExitCode = [int]$infoResult.ExitCode
    if ($infoResult.TimedOut) {
      Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
        status            = 'WingetInfoTimedOut'
        wingetExitCode    = $infoExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $infoExitCode
      } | Out-Null
      return $false
    }
    if ($infoExitCode -ne 0) {
      Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
        status            = 'WingetInfoFailed'
        wingetExitCode    = $infoExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $infoExitCode
      } | Out-Null
      return $false
    }
  }

  if (-not (Invoke-WingetSourceUpdate)) {
    if (-not (Invoke-WingetSourceReset)) {
      return $false
    }
  }

  return $true
}

function Get-PackageReleaseHold {
  param([object]$Package)
  $rule = Get-PackageRule -PackageId ([string]$Package.Id)
  # The deadline requests review; it must not authorize reinstalling unchanged known-bad bytes.
  $holds = @(Get-PropertyValue -Object $rule -Name 'releaseHolds' -Default @() | Where-Object { $_.version -eq $Package.Available -and $_.scope -eq (Get-WingetScopeArgument) })
  if ($holds.Count -eq 0) { return [pscustomobject]@{ Status = 'Clear' } }
  $arguments = @('show','--id',[string]$Package.Id,'--exact','--version',[string]$Package.Available,'--scope',(Get-WingetScopeArgument),'--accept-source-agreements','--disable-interactivity')
  if ($Package.Source) { $arguments += @('--source',[string]$Package.Source) }
  $result = Invoke-WingetProcess -Arguments $arguments -TimeoutSeconds 120
  $evidence = (Get-PackageAttemptEvidencePath -Package $Package -Attempt 0) -replace '\.attempt\.json$', '.release-hold.json'
  if (-not $WhatIfPreference) {
    [ordered]@{ packageId = $Package.Id; available = $Package.Available; arguments = $arguments; result = $result } | ConvertTo-Json -Depth 6 | Write-AtomicJson -Path $evidence
  }
  $hashes = @([regex]::Matches([string]$result.StdOut, '(?im)^\s*Installer SHA256:\s*([a-f0-9]{64})\s*$') | ForEach-Object { $_.Groups[1].Value })
  if ($result.TimedOut -or $result.ExitCode -ne 0 -or $hashes.Count -ne 1) { return [pscustomobject]@{ Status = 'DeferredReleaseHoldUnverified'; Reason = 'Cannot verify the candidate installer hash; no installation attempted.'; EvidencePath = $evidence } }
  $match = @($holds | Where-Object { $_.installerSha256 -eq $hashes[0] } | Select-Object -First 1)
  if ($match.Count) { return [pscustomobject]@{ Status = 'DeferredKnownBadRelease'; Reason = $match[0].reason; ReviewOverdue = ([datetime]$match[0].until -le (Get-Date)); ReviewBy = [string]$match[0].until; EvidencePath = $evidence } }
  return [pscustomobject]@{ Status = 'Clear' }
}

function Invoke-PackageUpgrade {
  param([object]$Package)

  if (-not (Test-MaintenanceWindow)) {
    Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{ packageId = [string]$Package.Id; status = 'DeferredOutsideMaintenanceWindow' } | Out-Null
    return 'Deferred'
  }
  $hold = Get-PackageReleaseHold -Package $Package
  if ($hold.Status -ne 'Clear') {
    if ($hold.Status -eq 'DeferredKnownBadRelease') { Move-FailureNotificationToAdminDeferral -Package $Package -Reason 'KnownBadRelease' -EvidencePath $hold.EvidencePath }
    Write-AgentEvent -EventId 5101 -EntryType Warning -Fields @{
      packageId = [string]$Package.Id; status = $hold.Status
      installed = [string]$Package.Installed; available = [string]$Package.Available
      source = [string]$Package.Source; requiresAdminRemediation = $true
      diagnostic = $hold.Reason; evidencePath = $hold.EvidencePath
      reviewOverdue = [bool](Get-PropertyValue $hold 'ReviewOverdue' $false)
      reviewBy = [string](Get-PropertyValue $hold 'ReviewBy' '')
    } | Out-Null
    return 'Deferred'
  }
  $existingDeferral = Get-InapplicableUpgradeDeferral -Package $Package
  if ($existingDeferral) {
    if ((Get-PropertyValue $existingDeferral 'reason' '') -in @('NoApplicableUpgrade', 'NoAvailableUpgrade', 'InstallerTechnologyMismatch', 'InstalledPackageNotFound', 'NoApplicableInstaller')) { $script:ApplicabilityDeferredCount++ }
    Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
      packageId     = [string]$Package.Id
      status        = 'Deferred' + (Get-PropertyValue -Object $existingDeferral -Name 'reason' -Default 'NoApplicableUpgrade') + 'Cached'
      installed     = [string]$Package.Installed
      available     = [string]$Package.Available
      source        = [string]$Package.Source
      deferredUntil = [string]$existingDeferral.deferredUntil
      evidencePath  = [string]$existingDeferral.evidencePath
      requiresAdminRemediation = (Get-PropertyValue -Object $existingDeferral -Name 'reason') -in @('InstallerTechnologyMismatch','InstalledPackageNotFound','NoApplicableInstaller')
    } | Out-Null
    return $(if ((Get-PropertyValue -Object $existingDeferral -Name 'reason') -eq 'InstallationOutcomeUncertain') { 'Failed' } else { 'Deferred' })
  }

  $maxAttempts = [int]$script:Policy.maxRetries + 1
  $recoveryActions = New-Object System.Collections.Generic.List[string]

  for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    if (-not (Test-MaintenanceWindow)) {
      Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{ packageId = [string]$Package.Id; status = 'DeferredOutsideMaintenanceWindow' } | Out-Null
      return 'Deferred'
    }
    $retryCount = $attempt - 1
    $evidencePath = Get-PackageAttemptEvidencePath -Package $Package -Attempt $attempt
    $packageLogPath = $evidencePath -replace '\.attempt\.json$', '.log'
    $verification = $null
    $args = @(
      'upgrade',
      '--id', [string]$Package.Id,
      '--exact',
      '--version', [string]$Package.Available,
      '--silent',
      '--disable-interactivity',
      '--accept-package-agreements',
      '--accept-source-agreements',
      '--scope', (Get-WingetScopeArgument),
      '--log', $packageLogPath
    )
    $rule = Get-PackageRule -PackageId ([string]$Package.Id)
    $installerType = if ($rule) { ([string]$rule.installerType).Trim() } else { '' }
    if ($installerType) { $args += @('--installer-type', $installerType) }
    if ($script:Policy.includeUnknown) { $args += '--include-unknown' }
    if ($Package.Source) { $args += @('--source', [string]$Package.Source) }

    $rebootBefore = Get-RebootMarkerSnapshot
    if ($PSCmdlet.ShouldProcess([string]$Package.Id, "winget upgrade attempt $attempt")) {
      $upgradeResult = Invoke-WingetProcess -Arguments $args -TimeoutSeconds 1800
      $wingetExitCode = [int]$upgradeResult.ExitCode
    } else {
      $wingetExitCode = 0
      $upgradeResult = [pscustomobject]@{ TimedOut = $false; ExitCode = 0; StdOut = ''; StdErr = '' }
    }

    $installerOutcome = Get-InstallerOutcomeFromLog -Path $packageLogPath
    $installerExitCode = $installerOutcome.ExitCode
    $commandOutcome = Get-WingetCommandOutcome -UpgradeResult $upgradeResult
    $rebootAfter = Get-RebootMarkerSnapshot
    Write-PackageAttemptEvidence -Package $Package -UpgradeResult $upgradeResult -InstallerExitCode $installerExitCode -InstallerOutcome $installerOutcome -CommandOutcome $commandOutcome -RebootBefore $rebootBefore -RebootAfter $rebootAfter -InstallerType $installerType -EvidencePath $evidencePath -Attempt $attempt -Arguments $args
    if ($upgradeResult.TimedOut) {
      $null = Set-InapplicableUpgradeDeferral -Package $Package -EvidencePath $evidencePath -CommandOutcome $commandOutcome
      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
        packageId         = [string]$Package.Id
        status            = 'InstallationOutcomeUncertain'
        failureReason     = 'InstallerMayStillBeRunningNoRetry'
        evidencePath      = $evidencePath
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($null -ne $installerExitCode -and $installerExitCode -notin @(0, 3010, 1641)) {
      if ($attempt -lt $maxAttempts -and (Test-RetryableInstallerExitCode -InstallerExitCode $installerExitCode)) {
        $recoveryActions.Add('retryInstallerFailure')
        Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
          packageId         = [string]$Package.Id
          status            = 'RetryScheduled'
          wingetExitCode    = $wingetExitCode
          wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
          installerExitCode = $installerExitCode
          retryCount        = $retryCount
          recoveryActions   = @($recoveryActions)
          logPath           = $packageLogPath
        } | Out-Null
        continue
      }

      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
        packageId         = [string]$Package.Id
        status            = 'Failed'
        failureReason     = if ($installerExitCode -eq 1603) { 'InstallerFatalNoRetry' } else { 'InstallerExitCodeNotRetryable' }
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($commandOutcome.Status -eq 'TransientDownloadFailure') {
      if ($attempt -lt $maxAttempts -and (Wait-DownloadRetry -Attempt $attempt)) {
        $recoveryActions.Add('retryTransientDownload')
        Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
          packageId         = [string]$Package.Id
          status            = 'TransientDownloadFailureRetryScheduled'
          wingetExitCode    = $wingetExitCode
          wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
          retryCount        = $retryCount
          recoveryActions   = @($recoveryActions)
          httpStatus        = $commandOutcome.HttpStatus
          hresult           = $commandOutcome.HResult
          downloadUrl       = $commandOutcome.DownloadUrl
          diagnostic        = $commandOutcome.Diagnostic
          evidencePath      = $evidencePath
          logPath           = $packageLogPath
        } | Out-Null
        continue
      }

      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
        packageId         = [string]$Package.Id
        status            = 'TransientDownloadFailure'
        failureReason     = 'AcquisitionFailureAfterRetryBudget'
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        httpStatus        = $commandOutcome.HttpStatus
        hresult           = $commandOutcome.HResult
        downloadUrl       = $commandOutcome.DownloadUrl
        diagnostic        = $commandOutcome.Diagnostic
        evidencePath      = $evidencePath
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($commandOutcome.Status -eq 'DownloadFailure') {
      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
        packageId         = [string]$Package.Id
        status            = 'DownloadFailure'
        failureReason     = 'DownloadFailureNotRetryable'
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        retryCount        = $retryCount
        httpStatus        = $commandOutcome.HttpStatus
        hresult           = $commandOutcome.HResult
        downloadUrl       = $commandOutcome.DownloadUrl
        diagnostic        = $commandOutcome.Diagnostic
        evidencePath      = $evidencePath
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($commandOutcome.Status -in @('NoApplicableUpgrade', 'NoAvailableUpgrade', 'InstallerTechnologyMismatch', 'InstalledPackageNotFound', 'NoApplicableInstaller')) {
      $script:ApplicabilityDeferredCount++
      $deferredUntil = Set-InapplicableUpgradeDeferral -Package $Package -EvidencePath $evidencePath -CommandOutcome $commandOutcome
      if ($commandOutcome.Status -eq 'NoApplicableInstaller') { Move-FailureNotificationToAdminDeferral -Package $Package -Reason 'NoApplicableInstaller' -EvidencePath $evidencePath }
      $adminRemediation = $commandOutcome.Status -in @('InstallerTechnologyMismatch','InstalledPackageNotFound','NoApplicableInstaller')
      Write-AgentEvent -EventId 5101 -EntryType $(if ($adminRemediation) { 'Warning' } else { 'Information' }) -Fields @{
        packageId     = [string]$Package.Id
        status        = 'Deferred' + $commandOutcome.Status
        requiresAdminRemediation = $adminRemediation
        scope         = Get-WingetScopeArgument
        installed     = [string]$Package.Installed
        available     = [string]$Package.Available
        source        = [string]$Package.Source
        deferredUntil = $deferredUntil.ToString('o')
        diagnostic    = $commandOutcome.Diagnostic
        evidencePath  = $evidencePath
        logPath       = $packageLogPath
      } | Out-Null
      return 'Deferred'
    }

    $restartEvidence = Test-RestartEvidence -InstallerExitCode $installerExitCode -UpgradeResult $upgradeResult -RebootBefore $rebootBefore -RebootAfter $rebootAfter
    if ($restartEvidence -and ($installerExitCode -in @(3010, 1641) -or $commandOutcome.Status -eq 'RestartRequired' -or ($wingetExitCode -in @(0, 3010, 1641) -and ($commandOutcome.Status -eq 'Succeeded' -or $installerOutcome.FinalSuccess)))) {
      Add-PendingRestartVerification -Package $Package -EvidencePath $evidencePath -Reason $restartEvidence
      Write-AgentEvent -EventId 5300 -EntryType Warning -Fields @{
        packageId         = [string]$Package.Id
        status            = 'RestartRequiredPendingVerification'
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        restartRequired   = $true
        logPath           = $packageLogPath
        evidencePath      = $evidencePath
        restartEvidence   = $restartEvidence
      } | Out-Null
      return 'RestartRequired'
    }

    if ($Mode -eq 'User' -and $wingetExitCode -eq 0 -and $commandOutcome.ApplicationRestartRequested) {
      Add-PendingApplicationRestart -Package $Package -EvidencePath $evidencePath
      Write-AgentEvent -EventId 5300 -EntryType Information -Fields @{
        packageId = $Package.Id; status = 'ApplicationRestartPendingVerification'; wingetExitCode = $wingetExitCode
        available = $Package.Available; source = $Package.Source; evidencePath = $evidencePath; logPath = $packageLogPath
      } | Out-Null
      return 'ApplicationRestartPending'
    }

    if ($wingetExitCode -in @(0, 3010, 1641) -and ($commandOutcome.Status -eq 'Succeeded' -or $installerOutcome.FinalSuccess)) {
      $verification = Test-WingetPackageStillOffered -Package $Package
      Write-PackageAttemptEvidence -Package $Package -UpgradeResult $upgradeResult -InstallerExitCode $installerExitCode -InstallerOutcome $installerOutcome -CommandOutcome $commandOutcome -RebootBefore $rebootBefore -RebootAfter $rebootAfter -InstallerType $installerType -EvidencePath $evidencePath -Attempt $attempt -Arguments $args -Verification $verification
      if ($verification.Status -ne 'VerifiedNotOffered') {
        if ($verification.Status -eq 'VerificationFailedStillOffered' -and $restartEvidence) {
          Add-PendingRestartVerification -Package $Package -EvidencePath $evidencePath -Reason $restartEvidence
          Write-AgentEvent -EventId 5300 -EntryType Warning -Fields @{
            packageId = [string]$Package.Id
            status = 'RestartRequiredPendingVerification'
            wingetExitCode = $wingetExitCode
            wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
            installerExitCode = $installerExitCode
            restartRequired = $true
            restartEvidence = $restartEvidence
            evidencePath = $evidencePath
            logPath = $packageLogPath
          } | Out-Null
          return 'RestartRequired'
        }
        Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
          packageId         = [string]$Package.Id
          status            = [string]$verification.Status
          wingetExitCode    = $wingetExitCode
          wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
          verificationExitCode = $verification.WingetExitCode
          verificationExitCodeHex = $verification.WingetExitCodeHex
          installerExitCode = $installerExitCode
          retryCount        = $retryCount
          recoveryActions   = @($recoveryActions)
          logPath           = $packageLogPath
          evidencePath      = $evidencePath
        } | Out-Null
        return 'Failed'
      }

      Write-AgentEvent -EventId 5100 -EntryType Information -Fields @{
        packageId         = [string]$Package.Id
        status            = 'Succeeded'
        wingetExitCode    = 0
        verificationStatus = [string]$verification.Status
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        logPath           = $packageLogPath
        evidencePath      = $evidencePath
      } | Out-Null
      Clear-InapplicableUpgradeDeferral -Package $Package
      return 'Succeeded'
    }

    if ($wingetExitCode -eq 0) {
      Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
        packageId = [string]$Package.Id
        status = 'UnconfirmedUpgradeOutcome'
        evidencePath = $evidencePath
        logPath = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    Write-AgentEvent -EventId 5400 -EntryType Error -Fields @{
      packageId         = [string]$Package.Id
      status            = 'Failed'
      wingetExitCode    = $wingetExitCode
      wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
      installerExitCode = $installerExitCode
      retryCount        = $retryCount
      recoveryActions   = @($recoveryActions)
      logPath           = $packageLogPath
    } | Out-Null
    return 'Failed'
  }
}

$script:EventLogUnavailable = $null
$script:Policy = New-DefaultPolicy

try {
  Initialize-AgentStorage
  if (-not (Enter-AgentRunLock)) {
    Write-AgentEvent -EventId 5001 -EntryType Information -Fields @{ status = 'SkippedConcurrentRun' } | Out-Null
    exit 0
  }
  Initialize-EventSource
  try { $script:Policy = Import-AgentPolicy } catch {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
      status = 'PolicyLoadFailedNoPatching'
      message = $_.Exception.Message
    } | Out-Null
    exit 2
  }
  $script:RunDeadline = (Get-Date).AddMinutes([int]$script:Policy.maxRunMinutes)
  Remove-OldLogs

  Write-AgentEvent -EventId 5000 -EntryType Information -Fields @{
    status = 'RunStarted'
  } | Out-Null
  Write-ActiveRunState

  if (-not $script:Policy.enabled) {
    Write-AgentEvent -EventId 5001 -EntryType Information -Fields @{
      status = 'DisabledByPolicy'
    } | Out-Null
    Clear-NotificationState -PreserveRestartRequired
    exit 0
  }

  if (-not (Test-MaintenanceWindow)) {
    Write-AgentEvent -EventId 5001 -EntryType Information -Fields @{
      status = 'SkippedOutsideMaintenanceWindow'
    } | Out-Null
    Clear-NotificationState -PreserveRestartRequired
    exit 0
  }

  if ($Mode -eq 'Machine' -and -not (Test-IsSystemAccount)) {
    Write-AgentEvent -EventId 5600 -EntryType Warning -Fields @{
      status = 'MachineModeNotRunningAsSystem'
    } | Out-Null
    Write-NotificationState -Status 'MachineModeNotRunningAsSystem' -Summary @{ Succeeded = 0; RestartRequired = 0; Failed = 1 } -RestartRequiredPackages @() -FailedPackages @()
    exit 2
  }

  if (-not (Test-WingetPreflight)) {
    Write-AgentEvent -EventId 5001 -EntryType Error -Fields @{
      status = 'PreflightFailed'
    } | Out-Null
    Write-NotificationState -Status 'PreflightFailed' -Summary @{ Succeeded = 0; RestartRequired = 0; Failed = 1 } -RestartRequiredPackages @() -FailedPackages @()
    exit 2
  }

  $inventory = Get-WingetInventory
  if ($inventory.Status -eq 'Failed') {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{ status = 'InventoryInconclusive'; message = $inventory.Message; evidencePath = $inventory.EvidencePath } | Out-Null
    exit 2
  }
  $availablePackages = @($inventory.Packages)
  Write-WingetInventorySnapshot -Packages $availablePackages
  $restartVerification = Resolve-PendingRestartVerification -AvailablePackages $availablePackages
  $reviewRestart = @($restartVerification.ReviewPackages)
  $applicationRestart = @(Resolve-PendingApplicationRestart -AvailablePackages $availablePackages)
  Resolve-RecordedPackageState -AvailablePackages $availablePackages -PostBootRestartPackages $restartVerification.PostBootPackages -ApplicationRestartPackages $applicationRestart
  $selectedPackages = @(Select-PolicyPackages -Packages $availablePackages -DeferredPackageIds $restartVerification.DeferredPackageIds -ApplicationRestartPackages $applicationRestart)

  $summary = @{
    Succeeded       = 0
    RestartRequired = 0
    Failed          = 0
    Deferred        = $script:SelectionDeferredCount
  }
  $restartRequiredPackages = New-Object System.Collections.Generic.List[object]
  $failedPackages = New-Object System.Collections.Generic.List[object]
  foreach ($pendingFailure in @($restartVerification.FailedPackages)) {
    if ($null -eq $pendingFailure) { continue }
    $summary.Failed++
    $failedPackages.Add($pendingFailure)
  }

  $newApplicationRestart = $false
  foreach ($pkg in $selectedPackages) {
    $result = Invoke-PackageUpgrade -Package $pkg
    if ($summary.ContainsKey($result)) {
      $summary[$result]++
    }
    if ($result -eq 'RestartRequired') {
      $restartRequiredPackages.Add($pkg)
    } elseif ($result -eq 'Failed') {
      $failedPackages.Add($pkg)
    } elseif ($result -eq 'ApplicationRestartPending') {
      $summary.Deferred++
      $newApplicationRestart = $true
    }
  }
  if ($newApplicationRestart) {
    $applicationRestart = @(Get-ApplicationRestartState)
    Resolve-RecordedPackageState -AvailablePackages $availablePackages -PostBootRestartPackages $restartVerification.PostBootPackages -ApplicationRestartPackages $applicationRestart
  }

  $rebootRequired = $summary.RestartRequired + $restartVerification.PendingRestartCount -gt $reviewRestart.Count
  $finalStatus = if ($summary.Failed -gt 0) { 'CompletedWithFailures' } elseif ($rebootRequired) { 'CompletedRestartRequired' } elseif ($reviewRestart.Count) { 'CompletedRestartVerificationInconclusive' } elseif ($selectedPackages.Count -eq 0) { 'NoPackagesSelected' } else { 'Completed' }
  $fields = Get-AssessmentSummaryFields -DiscoveredCount $availablePackages.Count -SelectedCount $selectedPackages.Count -Summary $summary -CarriedRestartCount $restartVerification.PendingRestartCount -PostRebootReviewCount $reviewRestart.Count -ApplicationRestartPendingCount $applicationRestart.Count
  $fields.status = $finalStatus
  if ($selectedPackages.Count -gt 0 -or $summary.Failed -gt 0) {
    $fields.packageId = '*'
    $fields.recoveryActions = @("Succeeded=$($summary.Succeeded)", "RestartRequired=$($summary.RestartRequired)", "Failed=$($summary.Failed)", "Deferred=$($summary.Deferred)")
  }
  Write-AgentEvent -EventId 5001 -EntryType Information -Fields $fields | Out-Null
  if ($summary.Failed -gt 0 -or $restartRequiredPackages.Count -gt 0) {
    Write-NotificationState -Status $finalStatus -Summary $summary -RestartRequiredPackages $restartRequiredPackages.ToArray() -FailedPackages $failedPackages.ToArray()
  } else {
    Clear-NotificationState -PreserveRestartRequired
  }

  if ($summary.Failed -gt 0) { exit 1 }
  exit 0
} catch {
  try {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
      status  = 'UnhandledAgentError'
      message = $_.Exception.Message
      scriptStackTrace = $_.ScriptStackTrace
    } | Out-Null
    Write-NotificationState -Status 'UnhandledAgentError' -Summary @{ Succeeded = 0; RestartRequired = 0; Failed = 1 } -RestartRequiredPackages @() -FailedPackages @()
  } catch {
    Write-Warning "Danmade Patch Agent failed before reporting was available: $($_.Exception.Message)"
  }
  exit 1
} finally {
  if ($null -ne $script:RunLock) {
    try { Clear-ActiveRunState } catch { }
    $script:RunLock.Dispose()
  }
}
