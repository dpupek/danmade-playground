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
    $process = Start-Process -FilePath $script:WingetCommand -ArgumentList $argumentList -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -PassThru
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
      '(?i)msi(?:\s+installer)?\s+(?:return|exit)\s+code\s*[:=]\s*(-?\d+)'
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
      Write-AgentEvent -EventId 5100 -EntryType Information -Fields @{
        packageId         = [string]$Package.Id
        status            = 'Succeeded'
        wingetExitCode    = 0
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
# MIIgCwYJKoZIhvcNAQcCoIIf/DCCH/gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgtOpO26J1LReX
# U1mqFJLza3r0aK7kgj6DyAPUXHrZpaCCGiYwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# ubdcMIIG6DCCBdCgAwIBAgITVgAAAMVxDkSgwgHrvgAAAAAAxTANBgkqhkiG9w0B
# AQsFADBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFzAVBgoJkiaJk/IsZAEZFgdu
# ZXhwb3J0MRYwFAYDVQQDEw1uZXhwb3J0LWxvY2FsMB4XDTI2MDYyNDE2NDYzMVoX
# DTI2MDcwNzIwMzEwMFowdDEVMBMGCgmSJomT8ixkARkWBWxvY2FsMRcwFQYKCZIm
# iZPyLGQBGRYHbmV4cG9ydDESMBAGA1UECxMJRGl2aXNpb25zMRowGAYDVQQLExFO
# ZXhwb3J0IFNvbHV0aW9uczESMBAGA1UEAxMJRGFuIFB1cGVrMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwneI0MpXdckSOPpsUq7m8qAr22igJfKu+B3K
# 4KD0+i49WATRD5XXGRqAijtuG+PSQmBPsSlXg737fEz0DZABRcaJn3oKQrzwXDY9
# GaprmDvALNG8AImTw2irrWIF2sl37TSZF3q5Bs45bR2U5pllJRABUYdbKZ+1Fh5k
# bm4g4U7zOO0MaVMib6Ye8fSxQQ8+DavbiW07ym1TWTFvcg4Owt7a9NCuHAeHvBDV
# PnWxrCcirwuAZL2n/JAhRRHpDAMnp0Cq01s9mp1foQAGSfiiJdN2t07QtZxTxpej
# L5eO2wEUc5+0jzibWkFy5iDXFzEx+UA/S/5twV3KBYDAnUad8QIDAQABo4IDnTCC
# A5kwPgYJKwYBBAGCNxUHBDEwLwYnKwYBBAGCNxUIhYmHZIX7mUqEuZsthev4V4f8
# lgOBFYT17Q2G2us2AgFkAgEAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB
# /wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBRv
# MnPZAnjFSFHO7U2J/OfFv8AIFjAfBgNVHSMEGDAWgBSn/qota2GbBcXK/qpf0af/
# GXhapDCBzQYDVR0fBIHFMIHCMIG/oIG8oIG5hoG2bGRhcDovLy9DTj1uZXhwb3J0
# LWxvY2FsLENOPU9LQ0RDMDIsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZp
# Y2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9bmV4cG9ydCxEQz1s
# b2NhbD9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9
# Y1JMRGlzdHJpYnV0aW9uUG9pbnQwggF/BggrBgEFBQcBAQSCAXEwggFtMIGuBggr
# BgEFBQcwAoaBoWxkYXA6Ly8vQ049bmV4cG9ydC1sb2NhbCxDTj1BSUEsQ049UHVi
# bGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlv
# bixEQz1uZXhwb3J0LERDPWxvY2FsP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RD
# bGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MFsGCCsGAQUFBzABhk9odHRwOi8v
# T0tDREMwMi5uZXhwb3J0LmxvY2FsL0NlcnRFbnJvbGwvT0tDREMwMi5uZXhwb3J0
# LmxvY2FsX25leHBvcnQtbG9jYWwuY3J0MF0GCCsGAQUFBzAChlFmaWxlOi8vLy9P
# S0NEQzAyLm5leHBvcnQubG9jYWwvQ2VydEVucm9sbC9PS0NEQzAyLm5leHBvcnQu
# bG9jYWxfbmV4cG9ydC1sb2NhbC5jcnQwMgYDVR0RBCswKaAnBgorBgEEAYI3FAID
# oBkMF2Rhbi5wdXBla0BuZXhwb3J0LmxvY2FsME4GCSsGAQQBgjcZAgRBMD+gPQYK
# KwYBBAGCNxkCAaAvBC1TLTEtNS0yMS0xMjU0MzE2MzQwLTI4NjQzMzE5NjgtNDc3
# OTA1NDQxLTExMDUwDQYJKoZIhvcNAQELBQADggEBAF3a4D/vdL/XqBZ2dtkiUrST
# sjJZjhuSQhl/yfxn/9CBYM131XTr8Q4eCYWk+Xz8txUFNM4khZ8IWttheF+jiWBn
# h4ROqtKSGJhgNV5xDgt7MrdBaltNg+kksDUnwR9Pxdzkke3ofQyL4M/TkZvCCUbQ
# 40eW4qAYhIrRgNBBeK8SC9+VpOtFgcmLbtBDeDNsu5VJXcvm30XhyqBXtSxpTmfm
# hcWTUlze80WL0orYtP2HZgjvShmjTWrXIt4Jl1TWVobvIs9LJ5h8qdqhUIGhEokb
# Bazmi//43AK8IZYHIpbR/0GEUF82VYQo6x9LLy/EmwiPw1X/8+X/hrl7+ZD7kXYw
# ggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHanlXRoMA0GCSqGSIb3DQEBCwUAMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcNMzYwOTAzMjM1OTU5WjBjMQswCQYD
# VQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lD
# ZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFtcCBSZXNwb25kZXIgMjAyNSAxMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0EasLRLGntDqrmBWsytXum9R
# /4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k
# +87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9
# A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXymOtRwJXcrcTTPPT2V1D/+cFllESvi
# H8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGH
# r7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD23DZgPfDrJJJK77epTwMP6eKA0kW
# a3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJukx7jphx40DQt82yepyekl4i0r8OEp
# s/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzix4A77p3awLbr89A90/nWGjXMGn7F
# QhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAeNIeWrzHKYueMJtItnj2Q+aTyLLKL
# M0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vMRHoRon4CWIvuiNN1Lk9Y+xZ66laz
# s2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOsNyEhzZtCGmnQigpFHti58CSmvEyJ
# cAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0O
# BBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8GA1UdIwQYMBaAFO9vU0rp5AZ8esri
# kFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIw
# MjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYy
# MDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJ
# KoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVv
# hREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FPsLSTwVQWo2H62yGBvg7ouCODwrx6
# ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0oU62PtgxOao872bOySCILdBghQ/Z
# LcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9lctZTTOJM3pHfKBAEcxQFoHlt2s9s
# XoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqI
# tH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs
# 7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E
# 5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGn
# oa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZ
# yvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVkT+um1vshETaWyQo8gmBto/m3acaP
# 9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvmpovq90K8eWyG2N01c4IhSOxqt81n
# MYIFOzCCBTcCAQEwXzBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFzAVBgoJkiaJ
# k/IsZAEZFgduZXhwb3J0MRYwFAYDVQQDEw1uZXhwb3J0LWxvY2FsAhNWAAAAxXEO
# RKDCAeu+AAAAAADFMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIDTQyaWT29m/hkB9LCui
# d/D9HwtydwCMxexOA+pF1++HMA0GCSqGSIb3DQEBAQUABIIBABTBGSq2bOnKh38h
# o9R+Gm7hrEn2z6+18UQfpnl8qmP9uyAU4ytUeY9WI6X+NlEZltvthVXSq6oX2nd5
# EnYKnasFgHuXeqUfLaY9Ic355SgVEiN2Jyo3yhZ/utgTpEANUO/MaAzVyLn1xdUB
# W0XLfeRkNTK2EJ/c+VDLgJIJd1xWPU/RgavePLV9BupKOdK10vR+gLcOddHtFQez
# xF2kbVtalgwZB+f2hg8+v95hxiAdHhLVWXVtGQBHlsICXcR5vQ/HUqGuheNAzB81
# /wqz5Dxb1u8tstUVNY+Y0nhINxIKYVHnV0KpF112uyELjuox2SPaJb75yges1TWZ
# 87Es0n6hggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# AhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkD
# MQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNjI5MTk0MzMyWjAvBgkq
# hkiG9w0BCQQxIgQg+KKHyCJK35lGocr3UVyY//q+H1m38m0lSrBRooiYD38wDQYJ
# KoZIhvcNAQEBBQAEggIABTDTDeONnGNRcNwExs14M3quObuduOQRDSRJEjOhBm4n
# A32De8oW7u8nTeYORf3VUWnxR3NAqE5BC+L+pgwkf3QFAl2Z/8wEO/bWX1id686J
# ln+J0FjechbaL6/PfybfWtahNqjqlzKT80pspS0V7iVEKzc0vjrN3YZUNUOpMq5c
# vTCyrdh1RpnC/Y+iaEIvtRwFQL9qEd7294sDG40VWpVuSRsAK/4Q3Vcpy446G/CW
# ypCOMtQDy8wljls2BuIdT43M8aP38N5vrO+XYKQRlQqR0ZjiSv7t9SYn/uScqk+b
# bugfziJ/8R3MLIbZtM5ElHqW5omBbEBsM/HpbHUfnvFpxUuk7Gkng6Pr9XB4dZyZ
# fyOuRrYQKLu2n0vr3oQdChiOPVK5h40KpH87ym4LvwuiZBXKG2I3syUkne7Iu9ew
# Aquizq2secdQ7h0SEGjovH8bYEWV38dlTbxM4UdYQgEPvKuALbHvEIZqTBeXOQO8
# 1GmdWFle7IJhMiR5XSxzTCKcP/dbyGRvB/48BJtxCIVPQ+tBTwVlCHn3HouC470z
# mgRXqZLujAMsfHElysn/ulwJM9wZNX4Bu9VMIee1b6/trbM9pGKMI4Yszu5qnVeT
# RccZHhQ4xpYez+CHc2ldPypgca/JsZkUeE+qxG02G51F8lJTWw+FmXsLInJOYcw=
# SIG # End signature block
