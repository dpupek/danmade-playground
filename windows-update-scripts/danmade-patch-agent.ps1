<#
.SYNOPSIS
Runs unattended winget package upgrades for domain-managed Windows endpoints.

.DESCRIPTION
Danmade Patch Agent is intended for Group Policy scheduled task deployment.
It uses a domain-distributed JSON policy when available, falls back to safe
defaults, emits Wazuh-friendly Event Log and JSONL records, and performs only
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
  '{0}-{1}-{2}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $env:COMPUTERNAME, $Mode
} else {
  $RunId
}
$script:Policy = $null
$script:JsonlPath = $null
$script:WingetCommand = $null
$script:CachedUninstallEntries = $null

function New-DefaultPolicy {
  [pscustomobject]@{
    enabled           = $true
    includeUnknown    = $true
    allowedPackageIds = @()
    blockedPackageIds = @()
    maxRetries        = 2
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

  return [pscustomobject]@{
    enabled           = [bool](Get-PropertyValue -Object $LoadedPolicy -Name 'enabled' -Default $default.enabled)
    includeUnknown    = [bool](Get-PropertyValue -Object $LoadedPolicy -Name 'includeUnknown' -Default $default.includeUnknown)
    allowedPackageIds = Convert-ToStringArray -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'allowedPackageIds' -Default @())
    blockedPackageIds = Convert-ToStringArray -Value (Get-PropertyValue -Object $LoadedPolicy -Name 'blockedPackageIds' -Default @())
    maxRetries        = $maxRetries
    maintenanceWindow = [pscustomobject]@{
      enabled = [bool](Get-PropertyValue -Object $maintenance -Name 'enabled' -Default $default.maintenanceWindow.enabled)
      start   = [string](Get-PropertyValue -Object $maintenance -Name 'start' -Default $default.maintenanceWindow.start)
      end     = [string](Get-PropertyValue -Object $maintenance -Name 'end' -Default $default.maintenanceWindow.end)
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
    return Merge-Policy -LoadedPolicy $null
  }

  try {
    $json = Get-Content -LiteralPath $resolvedPath -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($json)) {
      throw "Policy file is empty: $resolvedPath"
    }
    return Merge-Policy -LoadedPolicy ($json | ConvertFrom-Json -ErrorAction Stop)
  } catch {
    $script:PolicyLoadError = $_.Exception.Message
    return Merge-Policy -LoadedPolicy $null
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

  try {
    # Child winget console processes must not flash in the interactive user task.
    $process = Start-Process -FilePath $script:WingetCommand -ArgumentList $argumentList -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
      try { $process.Kill() } catch { }
      return [pscustomobject]@{
        ExitCode = -1
        TimedOut = $true
        StdOut   = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw } else { '' }
        StdErr   = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
      }
    }

    return [pscustomobject]@{
      ExitCode = [int]$process.ExitCode
      TimedOut = $false
      StdOut   = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw } else { '' }
      StdErr   = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
    }
  } finally {
    Remove-Item -LiteralPath $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
  }
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
  $script:JsonlPath = Join-Path -Path $LogRoot -ChildPath 'Events\patch-agent.jsonl'
}

function Get-NotificationStatePath {
  return (Join-Path -Path $LogRoot -ChildPath 'State\pending-notification.json')
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
  $succeededCount = 0
  $restartRequiredCount = 0
  $failedCount = 0
  if ($Summary.ContainsKey('Succeeded')) { $succeededCount = [int]$Summary.Succeeded }
  if ($Summary.ContainsKey('RestartRequired')) { $restartRequiredCount = [int]$Summary.RestartRequired }
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
      restartRequired = $restartRequiredCount
      failed          = $failedCount
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
    $state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $statePath -Encoding utf8
  }
}

function Clear-NotificationState {
  param([switch]$PreserveRestartRequired)

  $statePath = Get-NotificationStatePath
  if (-not (Test-Path -LiteralPath $statePath)) { return }
  if ($PreserveRestartRequired) {
    try {
      $state = Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      if ($state.PSObject.Properties['status'] -and [string]$state.status -eq 'CompletedRestartRequired') {
        return
      }
    } catch {
      # If state is unreadable, clear it so a corrupt notification does not keep nagging users.
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
  foreach ($path in @((Join-Path $LogRoot 'Logs'), (Join-Path $LogRoot 'Events'))) {
    if (-not (Test-Path -LiteralPath $path)) { continue }
    Get-ChildItem -LiteralPath $path -File -ErrorAction SilentlyContinue |
      Where-Object { $_.LastWriteTime -lt $cutoff } |
      Remove-Item -Force -ErrorAction SilentlyContinue
  }
}

function Test-MaintenanceWindow {
  $window = $script:Policy.maintenanceWindow
  if (-not $window.enabled) { return $true }

  $start = [TimeSpan]::Zero
  $end = [TimeSpan]::Zero
  if (-not [TimeSpan]::TryParse([string]$window.start, [ref]$start)) { return $true }
  if (-not [TimeSpan]::TryParse([string]$window.end, [ref]$end)) { return $true }

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

function Resolve-PackageInstallScope {
  param([object]$Package)

  if (-not $Package -or [string]::IsNullOrWhiteSpace([string]$Package.Name)) {
    return 'Unknown'
  }

  $packageNameNorm = Normalize-DisplayText -Text ([string]$Package.Name)
  if (-not $packageNameNorm) { return 'Unknown' }

  $bestMatch = $null
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

    if ($score -gt $bestScore) {
      $bestScore = $score
      $bestMatch = $entry
    }
  }

  if (-not $bestMatch) { return 'Unknown' }
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

function Get-WingetUpgradeList {
  $baseArgs = @('upgrade', '--disable-interactivity', '--accept-source-agreements')
  if ($script:Policy.includeUnknown) { $baseArgs += '--include-unknown' }

  try {
    $jsonResult = Invoke-WingetProcess -Arguments ($baseArgs + @('--output', 'json')) -TimeoutSeconds 180
    if ($jsonResult.TimedOut) {
      Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
        status            = 'WingetUpgradeListJsonTimedOut'
        wingetExitCode    = $jsonResult.ExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $jsonResult.ExitCode
      } | Out-Null
      throw 'winget JSON upgrade listing timed out'
    }
    $json = $jsonResult.StdOut
    if ($json) {
      $jsonText = ([string]$json).Trim()
      $firstBraceIndex = $jsonText.IndexOfAny(@('[', '{'))
      if ($firstBraceIndex -ge 0) {
        $data = $jsonText.Substring($firstBraceIndex) | ConvertFrom-Json -ErrorAction Stop
        $collected = New-Object System.Collections.Generic.List[object]
        Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null
        if ($collected.Count -gt 0) {
          $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
          return @(
            foreach ($pkg in $collected) {
              if (-not $pkg.Id) { continue }
              if ($seen.Add($pkg.Id)) { Add-PackageScope -Package $pkg }
            }
          )
        }
      }
    }
  } catch {
    Write-Verbose "winget JSON upgrade listing failed; using table fallback. $($_.Exception.Message)"
  }

  try {
    $tableResult = Invoke-WingetProcess -Arguments $baseArgs -TimeoutSeconds 180
    if ($tableResult.TimedOut) {
      Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
        status            = 'WingetUpgradeListTableTimedOut'
        wingetExitCode    = $tableResult.ExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $tableResult.ExitCode
      } | Out-Null
      return @()
    }
    $output = $tableResult.StdOut
    if (-not $output) { return @() }

    $lines = $output -split "`r`n|`n|`r"
    $header = $lines | Where-Object { $_ -match '^Name\s+Id\s+Version\s+Available\s+Source' } | Select-Object -First 1
    if (-not $header) { return @() }

    $headerIndex = [array]::IndexOf($lines, $header)
    $candidateLines = $lines | Select-Object -Skip ($headerIndex + 2)
    return @(
      foreach ($line in $candidateLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '^[-\s]+$') { continue }
        if ($line -match '^\d+\s+upgrades available\.?$') { continue }
        if ($line -match '^The following packages have an upgrade available') { continue }
        $pkg = ConvertFrom-WingetUpgradeTableRow -Line $line
        if ($pkg) { Add-PackageScope -Package $pkg }
      }
    )
  } catch {
    Write-Verbose "winget table upgrade listing failed. $($_.Exception.Message)"
    return @()
  }
}

function Select-PolicyPackages {
  param([object[]]$Packages)

  $allowed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($id in @($script:Policy.allowedPackageIds)) { [void]$allowed.Add($id) }

  $blocked = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($id in @($script:Policy.blockedPackageIds)) { [void]$blocked.Add($id) }

  return @(
    foreach ($pkg in @($Packages)) {
      if (-not $pkg.Id) { continue }
      if ($blocked.Contains([string]$pkg.Id)) {
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedBlocked'
        } | Out-Null
        continue
      }
      if ($allowed.Count -gt 0 -and -not $allowed.Contains([string]$pkg.Id)) {
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedNotAllowed'
        } | Out-Null
        continue
      }
      $scope = if ($pkg.PSObject.Properties['InstallScope']) { [string]$pkg.InstallScope } else { 'Unknown' }
      if ($Mode -eq 'User' -and $scope -ne 'User') {
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedScope'
          installScope = $scope
        } | Out-Null
        continue
      }
      if ($Mode -eq 'Machine' -and $scope -eq 'User') {
        Write-AgentEvent -EventId 5101 -EntryType Information -Fields @{
          packageId = [string]$pkg.Id
          status    = 'SkippedScope'
          installScope = $scope
        } | Out-Null
        continue
      }
      $pkg
    }
  )
}

function Get-InstallerExitCodeFromLog {
  param([string]$Path)
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }

  try {
    $text = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $patterns = @(
      '(?i)(?:install|uninstall)\s+failed\s+with\s+exit\s+code:\s*(-?\d+)',
      '(?i)installer\s+return\s+code\s*[:=]\s*(-?\d+)',
      '(?i)msi(?:\s+installer)?\s+(?:return|exit)\s+code\s*[:=]\s*(-?\d+)',
      '(?i)returned\s+actual\s+error\s+code\s+(-?\d+)',
      '(?i)errorcode\s*:\s*(-?\d+)'
    )
    foreach ($pattern in $patterns) {
      $match = [regex]::Match($text, $pattern)
      if (-not $match.Success) { continue }
      $parsed = 0
      if ([int]::TryParse($match.Groups[1].Value, [ref]$parsed)) { return $parsed }
    }
  } catch {
    return $null
  }

  return $null
}

function Test-WingetPackageStillOffered {
  param([object]$Package)

  $args = @(
    'upgrade',
    '--id', [string]$Package.Id,
    '--exact',
    '--disable-interactivity',
    '--accept-source-agreements',
    '--scope', (Get-WingetScopeArgument)
  )
  if ($script:Policy.includeUnknown) { $args += '--include-unknown' }
  if ($Package.Source) { $args += @('--source', [string]$Package.Source) }

  $result = Invoke-WingetProcess -Arguments $args -TimeoutSeconds 180
  $exitCode = [int]$result.ExitCode
  if ($result.TimedOut) {
    return [pscustomobject]@{
      Status            = 'VerificationInconclusive'
      StillOffered      = $false
      TimedOut          = $true
      WingetExitCode    = $exitCode
      WingetExitCodeHex = Convert-ToHexCode -Code $exitCode
    }
  }

  if ($exitCode -ne 0) {
    return [pscustomobject]@{
      Status            = 'VerificationInconclusive'
      StillOffered      = $false
      TimedOut          = $false
      WingetExitCode    = $exitCode
      WingetExitCodeHex = Convert-ToHexCode -Code $exitCode
    }
  }

  $packageId = [string]$Package.Id
  $stillOffered = $false
  $stdout = [string]$result.StdOut
  if ($stdout) {
    $jsonText = $stdout.Trim()
    $firstJsonIndex = $jsonText.IndexOfAny(@('[', '{'))
    if ($firstJsonIndex -ge 0) {
      try {
        $data = $jsonText.Substring($firstJsonIndex) | ConvertFrom-Json -ErrorAction Stop
        $collected = New-Object System.Collections.Generic.List[object]
        Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null
        $stillOffered = @($collected | Where-Object { $_.Id -and [string]::Equals([string]$_.Id, $packageId, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
      } catch {
        $stillOffered = $false
      }
    }

    if (-not $stillOffered) {
      foreach ($line in @($stdout -split "`r`n|`n|`r")) {
        $pkg = ConvertFrom-WingetUpgradeTableRow -Line $line
        if ($pkg -and [string]::Equals([string]$pkg.Id, $packageId, [System.StringComparison]::OrdinalIgnoreCase)) {
          $stillOffered = $true
          break
        }
      }
    }
  }

  if ($stillOffered) {
    return [pscustomobject]@{
      Status            = 'VerificationFailedStillOffered'
      StillOffered      = $true
      TimedOut          = $false
      WingetExitCode    = $exitCode
      WingetExitCodeHex = Convert-ToHexCode -Code $exitCode
    }
  }

  return [pscustomobject]@{
    Status            = 'VerifiedNotOffered'
    StillOffered      = $false
    TimedOut          = $false
    WingetExitCode    = $exitCode
    WingetExitCodeHex = Convert-ToHexCode -Code $exitCode
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

function Invoke-PackageUpgrade {
  param([object]$Package)

  $safeId = ([string]$Package.Id -replace '[^A-Za-z0-9._-]', '_')
  $packageLogPath = Join-Path -Path (Join-Path -Path $LogRoot -ChildPath 'Logs') -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))
  $maxAttempts = [int]$script:Policy.maxRetries + 1
  $recoveryActions = New-Object System.Collections.Generic.List[string]

  for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    $retryCount = $attempt - 1
    $args = @(
      'upgrade',
      '--id', [string]$Package.Id,
      '--exact',
      '--silent',
      '--disable-interactivity',
      '--accept-package-agreements',
      '--accept-source-agreements',
      '--scope', (Get-WingetScopeArgument),
      '--log', $packageLogPath
    )
    if ($script:Policy.includeUnknown) { $args += '--include-unknown' }
    if ($Package.Source) { $args += @('--source', [string]$Package.Source) }

    if ($PSCmdlet.ShouldProcess([string]$Package.Id, "winget upgrade attempt $attempt")) {
      $upgradeResult = Invoke-WingetProcess -Arguments $args -TimeoutSeconds 1800
      $wingetExitCode = [int]$upgradeResult.ExitCode
    } else {
      $wingetExitCode = 0
      $upgradeResult = [pscustomobject]@{ TimedOut = $false }
    }

    $installerExitCode = Get-InstallerExitCodeFromLog -Path $packageLogPath
    if ($upgradeResult.TimedOut) {
      if ($attempt -lt $maxAttempts) {
        $recoveryActions.Add('retryAfterTimeout')
        Write-AgentEvent -EventId 5200 -EntryType Warning -Fields @{
          packageId         = [string]$Package.Id
          status            = 'PackageUpgradeTimedOutRetryScheduled'
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
        status            = 'PackageUpgradeTimedOut'
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($installerExitCode -eq 3010) {
      Write-AgentEvent -EventId 5300 -EntryType Warning -Fields @{
        packageId         = [string]$Package.Id
        status            = 'RestartRequired'
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        restartRequired   = $true
        logPath           = $packageLogPath
      } | Out-Null
      return 'RestartRequired'
    }

    if ($null -ne $installerExitCode -and $installerExitCode -ne 0) {
      if ($attempt -lt $maxAttempts) {
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
        wingetExitCode    = $wingetExitCode
        wingetExitCodeHex = Convert-ToHexCode -Code $wingetExitCode
        installerExitCode = $installerExitCode
        retryCount        = $retryCount
        recoveryActions   = @($recoveryActions)
        logPath           = $packageLogPath
      } | Out-Null
      return 'Failed'
    }

    if ($wingetExitCode -eq 0) {
      $verification = Test-WingetPackageStillOffered -Package $Package
      if ($verification.Status -ne 'VerifiedNotOffered') {
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
      } | Out-Null
      return 'Succeeded'
    }

    if ($attempt -lt $maxAttempts) {
      $recoveryActions.Add('retrySilent')
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

$script:PolicyLoadError = $null
$script:EventLogUnavailable = $null
$script:Policy = Import-AgentPolicy

try {
  Initialize-AgentStorage
  Initialize-EventSource
  Remove-OldLogs

  if ($script:PolicyLoadError) {
    Write-AgentEvent -EventId 5600 -EntryType Warning -Fields @{
      status = 'PolicyLoadFailedUsingDefaults'
      message = $script:PolicyLoadError
    } | Out-Null
  }

  Write-AgentEvent -EventId 5000 -EntryType Information -Fields @{
    status = 'RunStarted'
  } | Out-Null

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

  $availablePackages = @(Get-WingetUpgradeList)
  $selectedPackages = @(Select-PolicyPackages -Packages $availablePackages)

  if ($selectedPackages.Count -eq 0) {
    Write-AgentEvent -EventId 5001 -EntryType Information -Fields @{
      status = 'NoPackagesSelected'
    } | Out-Null
    Clear-NotificationState -PreserveRestartRequired
    exit 0
  }

  $summary = @{
    Succeeded       = 0
    RestartRequired = 0
    Failed          = 0
  }
  $restartRequiredPackages = New-Object System.Collections.Generic.List[object]
  $failedPackages = New-Object System.Collections.Generic.List[object]

  foreach ($pkg in $selectedPackages) {
    $result = Invoke-PackageUpgrade -Package $pkg
    if ($summary.ContainsKey($result)) {
      $summary[$result]++
    }
    if ($result -eq 'RestartRequired') {
      $restartRequiredPackages.Add($pkg)
    } elseif ($result -eq 'Failed') {
      $failedPackages.Add($pkg)
    }
  }

  $finalStatus = if ($summary.Failed -gt 0) { 'CompletedWithFailures' } elseif ($summary.RestartRequired -gt 0) { 'CompletedRestartRequired' } else { 'Completed' }
  Write-AgentEvent -EventId 5001 -EntryType Information -Fields @{
    status          = $finalStatus
    packageId       = '*'
    recoveryActions = @("Succeeded=$($summary.Succeeded)", "RestartRequired=$($summary.RestartRequired)", "Failed=$($summary.Failed)")
  } | Out-Null
  if ($finalStatus -eq 'Completed') {
    Clear-NotificationState
  } else {
    Write-NotificationState -Status $finalStatus -Summary $summary -RestartRequiredPackages $restartRequiredPackages.ToArray() -FailedPackages $failedPackages.ToArray()
  }

  if ($summary.Failed -gt 0) { exit 1 }
  exit 0
} catch {
  try {
    Write-AgentEvent -EventId 5600 -EntryType Error -Fields @{
      status  = 'UnhandledAgentError'
      message = $_.Exception.Message
    } | Out-Null
    Write-NotificationState -Status 'UnhandledAgentError' -Summary @{ Succeeded = 0; RestartRequired = 0; Failed = 1 } -RestartRequiredPackages @() -FailedPackages @()
  } catch {
    Write-Warning "Danmade Patch Agent failed before reporting was available: $($_.Exception.Message)"
  }
  exit 1
}

# SIG # Begin signature block
# MIIgEQYJKoZIhvcNAQcCoIIgAjCCH/4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH44Y53X2Z/QKZ
# b7HuKarlLTaHGictXHVjHLovzCU0XKCCGiwwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwgga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphB
# cr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6p
# vF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHe
# HYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEd
# gkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjU
# jsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bR
# VFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeS
# LsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIV
# NSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL
# 6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2Zd
# SoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFU
# eEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/
# BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0j
# BBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8E
# PDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEw
# DQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/
# T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQ
# E7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9r
# EVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y
# 1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gx
# dEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3t
# y9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcy
# tL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEB
# YTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud
# /v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiS
# uEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZP
# ubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC0cR2p5V0aDANBgkqhkiG9w0BAQsF
# ADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNV
# BAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hB
# MjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAwMFoXDTM2MDkwMzIzNTk1OVowYzEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJE
# aWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjUg
# MTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANBGrC0Sxp7Q6q5gVrMr
# V7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA69HFTBdwbHwBSOeLpvPnZ8ZN+vo8
# dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6wW2R6kSu9RJt/4QhguSssp3qome7M
# rxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00Cll8pjrUcCV3K3E0zz09ldQ//nBZ
# ZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOMA3CoB/iUSROUINDT98oksouTMYFO
# nHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmotuQhcg9tw2YD3w6ySSSu+3qU8DD+n
# igNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1OpbybpMe46YceNA0LfNsnqcnpJeIt
# K/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeHVZlc4seAO+6d2sC26/PQPdP51ho1
# zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1roSrgHjSHlq8xymLnjCbSLZ49kPmk
# 8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSUROwnu7zER6EaJ+AliL7ojTdS5PWPsW
# eupWs7NpChUk555K096V1hE0yZIXe+giAwW00aHzrDchIc2bQhpp0IoKRR7YufAk
# prxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0G
# A1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM6DAfBgNVHSMEGDAWgBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEy
# NTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hB
# MjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcEua5gQezRCESeY0ByIfjk9iJP2zWL
# pQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/YmRDfxT7C0k8FUFqNh+tshgb4O6Lgj
# g8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8AQ/UdKFOtj7YMTmqPO9mzskgiC3Q
# YIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/EABgfZXLWU0ziTN6R3ygQBHMUBaB5
# bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQVTeLni2nHkX/QqvXnNb+YkDFkxUG
# tMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gVutDojBIFeRlqAcuEVT0cKsb+zJNE
# suEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85EE8LUkqRhoS3Y50OHgaY7T/lwd6U
# Arb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hggt8l2Yv7roancJIFcbojBcxlRcGG
# 0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJgKf47CdxVRd/ndUlQ05oxYy2zRWV
# FjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLvUxxVZE/rptb7IRE2lskKPIJgbaP5
# t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7POGT75qaL6vdCvHlshtjdNXOCIUjs
# arfNZzCCBu4wggXWoAMCAQICE1YAAAD7uI/G5fXiwIQAAQAAAPswDQYJKoZIhvcN
# AQELBQAwSDEVMBMGCgmSJomT8ixkARkWBWxvY2FsMRcwFQYKCZImiZPyLGQBGRYH
# bmV4cG9ydDEWMBQGA1UEAxMNbmV4cG9ydC1sb2NhbDAeFw0yNjA3MDYxNjEzMjla
# Fw0yNzA3MDYxNjEzMjlaMHQxFTATBgoJkiaJk/IsZAEZFgVsb2NhbDEXMBUGCgmS
# JomT8ixkARkWB25leHBvcnQxEjAQBgNVBAsTCURpdmlzaW9uczEaMBgGA1UECxMR
# TmV4cG9ydCBTb2x1dGlvbnMxEjAQBgNVBAMTCURhbiBQdXBlazCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAMZQ4zJZc+eugIuUhFyk+eVYC9gkueExn2Sx
# JvuJ+K9v9ykS2jTErJWxVzACQokpRuu7X7t7ZTkMIjQvl7ovzxDsRHOZ93gvBR61
# D7rwHUll1lMqOUzoA/5pHRMmlrxzLk0TK5I7XW0aCU94RHfktU0DIb0MpiGUyKcO
# 5LCVer4D0lXzqTTIDiizY+lJPRYKNfxnK3tGdROO/NObmWeFOlKnDlNxmCBsSMv8
# e6gBs7pk70TTUAf6yq3sCQ45Oo+q58Z0dMw5+5pir2K3pY/En/BWsULa5Qn9Xy66
# lcr19YpqVjq83NoOtgmaw4hGz8zjGdZLOj9YA3bKQNxDDZqcCqkCAwEAAaOCA6Mw
# ggOfMD4GCSsGAQQBgjcVBwQxMC8GJysGAQQBgjcVCIWJh2SF+5lKhLmbLYXr+FeH
# /JYDgRWE9e0NhtrrNgIBZAIBADATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNVHQ8B
# Af8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU
# egRE7DA1bfxWX5Xw8IpDWAn/P3wwHwYDVR0jBBgwFoAUp/6qLWthmwXFyv6qX9Gn
# /xl4WqQwgc0GA1UdHwSBxTCBwjCBv6CBvKCBuYaBtmxkYXA6Ly8vQ049bmV4cG9y
# dC1sb2NhbCxDTj1PS0NEQzAyLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2
# aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPW5leHBvcnQsREM9
# bG9jYWw/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNz
# PWNSTERpc3RyaWJ1dGlvblBvaW50MIIBhQYIKwYBBQUHAQEEggF3MIIBczCBrgYI
# KwYBBQUHMAKGgaFsZGFwOi8vL0NOPW5leHBvcnQtbG9jYWwsQ049QUlBLENOPVB1
# YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRp
# b24sREM9bmV4cG9ydCxEQz1sb2NhbD9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0
# Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTBeBggrBgEFBQcwAYZSaHR0cDov
# L09LQ0RDMDIubmV4cG9ydC5sb2NhbC9DZXJ0RW5yb2xsL09LQ0RDMDIubmV4cG9y
# dC5sb2NhbF9uZXhwb3J0LWxvY2FsKDEpLmNydDBgBggrBgEFBQcwAoZUZmlsZTov
# Ly8vT0tDREMwMi5uZXhwb3J0LmxvY2FsL0NlcnRFbnJvbGwvT0tDREMwMi5uZXhw
# b3J0LmxvY2FsX25leHBvcnQtbG9jYWwoMSkuY3J0MDIGA1UdEQQrMCmgJwYKKwYB
# BAGCNxQCA6AZDBdkYW4ucHVwZWtAbmV4cG9ydC5sb2NhbDBOBgkrBgEEAYI3GQIE
# QTA/oD0GCisGAQQBgjcZAgGgLwQtUy0xLTUtMjEtMTI1NDMxNjM0MC0yODY0MzMx
# OTY4LTQ3NzkwNTQ0MS0xMTA1MA0GCSqGSIb3DQEBCwUAA4IBAQCbq973SpW2d0oM
# JAyWvJiZWDyuPCmzSubthWPRxVhbi6vdh40+rt3fZ6uTJ2Z8MdGLqtmNPK1HXBuj
# 0X0/2m/5MUzyYmZjfqeRyBuSzXAdfWuZ27J5jhOOj51+KDAPEOxjRea6MZBr1ezC
# K7asWYX1mcK19E2ZlnZTic0fx2cCVJNRTTNaQmYvxnJq0XNbbbaZcOSIW8TTERsb
# p+NCKwm0mzfrIvA3QaX0DrpYTdNLgwm0u7kD7d/f9S/aGbXZp6L/VUqNPWc9Ld0x
# OAu8qXJ/zFjIbs4MeClYeLsWTn0f6h3+J++egix4lZZTBfmn4tcPJL2QT7WzVKPG
# +lmwH+uqMYIFOzCCBTcCAQEwXzBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFzAV
# BgoJkiaJk/IsZAEZFgduZXhwb3J0MRYwFAYDVQQDEw1uZXhwb3J0LWxvY2FsAhNW
# AAAA+7iPxuX14sCEAAEAAAD7MA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIP2xsC7AbRXz
# oqu3RrmRVi9VltfrUe1a7yqqHWwrUYkJMA0GCSqGSIb3DQEBAQUABIIBAEm5vGPK
# QToHrUo5EOPHnNqg0uWAH7Y/JKEjBa6s87oAMstjaHXIHYUl7mWFeLe9wSfW+TaK
# XF+LDtkzPepnpVc9KvX/tI31zgrhy+LyeSxg6Wix7YpnjZfd14KlySqCxTWyprKX
# epJJZIa1F2Zjd800xztbedxcPjrh5MH95/8ALwcc5N5xPhLaM+nCPO1+y44aGjrF
# llW3R18EFwsK+lzjCmnUGZIEguXu33bVVcZo4lS4XbpkBMqwQPYIGWRBHx/VS5+I
# AkN09fIYoDjpiu96vFLY2t/Bff0MrogPmAdxDeiA73KBlA+/Cwj04MZ/cs8+qDkj
# K+4bAAYhSsGkCqShggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNzEwMTI0NzAw
# WjAvBgkqhkiG9w0BCQQxIgQg1cKHYli6pcJZTbUartBtqjxS7G6l0DWH06CJp315
# 3KcwDQYJKoZIhvcNAQEBBQAEggIAED7gkkkFxkxEF1KygCHdisZF/T8WxyHvsvs6
# gXGrjSEA2QRSHpfHshg/gkRxai0nJ/2PmCNYh40Ekjb+Hb+uF2l2Gda1Ho57GopH
# sc+tB2qh8HVzujKV+ZhBRJi2eLxkHnmhPT/dMPPXPgIFuFSc4x2tSRpW6Tce/kV4
# VuPo9rY6C/L7thRYMdcS6zdP3ykGJl6XOt+u/nVbYr4tYPmu3/3tSwZtB15zPNC1
# bo4YS1jO4O+hNKD36ZRoCzlGoF9Jg69/VvwB8rzDy5MJZoAdYbJ1z6B2eKCVTp0d
# NzHkYy4tsiE6Y1KJ2Q1q4e8PTG7hP8IHjGq6y2Ka5w6/ltbKoq5OJQCb69rglPJw
# vUdd/VvVIUtXP7nkf6KkNc2H3QOCfcgqQD4WXsrOyuFxluaLvgicz8pbzd1ixY1o
# yblXBReQFC52WfdTlVqrlqLFCRpu0BVjVIrYoOQU8o51/bDUClVrRrnSGjeIR0Xg
# gS0dWGz24AB1x8mRqDLc2hmVizr1ZJZNCuyY3yrxfPt4C8JsbX7XkPkPWjSZMw4A
# nACngpynRv1YimcYgYQWcOj3amKjr2p0SircrUUxvXvPKO4jBR4x0M1yYaUESs5J
# STSN8jR0U9uygfCSJt/KdgZMp4m77wqcTeh6m5utRCnMPnqdcIh/YJI8pqWQt0UT
# OUiNoYY=
# SIG # End signature block
