<#
.SYNOPSIS
Reports stale Danmade Patch Agent machine processes without changing endpoint state.

.DESCRIPTION
Reads the Danmade Patch Agent policy for maxRunMinutes, inspects only matching
PowerShell processes, and records stale-process evidence in the shared agent
JSONL log. It does not stop processes, disable scheduled tasks, or modify other
scheduled tasks. ProcessSnapshotPath exists only for deterministic local tests.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
  [string]$PolicyPath,

  [string]$LogRoot = 'C:\ProgramData\DanmadePatchAgent',

  [string]$ProcessSnapshotPath,

  [string]$TaskSnapshotPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PropertyValue {
  param([object]$Object, [string]$Name, [object]$Default = $null)

  if ($null -eq $Object) { return $Default }
  $property = $Object.PSObject.Properties[$Name]
  if ($null -eq $property -or $null -eq $property.Value) { return $Default }
  return $property.Value
}

function Get-PolicyMaxRunMinutes {
  $maxRunMinutes = 120
  if (-not [string]::IsNullOrWhiteSpace($PolicyPath) -and (Test-Path -LiteralPath $PolicyPath)) {
    try {
      $policy = Get-Content -LiteralPath $PolicyPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      $maxRunMinutes = [int](Get-PropertyValue -Object $policy -Name 'maxRunMinutes' -Default $maxRunMinutes)
    } catch {
      $maxRunMinutes = 120
    }
  }
  if ($maxRunMinutes -lt 15) { return 15 }
  if ($maxRunMinutes -gt 240) { return 240 }
  return $maxRunMinutes
}

function Get-ActiveRunState {
  $path = Join-Path -Path $LogRoot -ChildPath 'State\active-run.json'
  if (-not (Test-Path -LiteralPath $path)) { return $null }
  try {
    $state = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if (-not (Get-PropertyValue $state 'runId') -or [int](Get-PropertyValue $state 'processId' 0) -le 0) { throw 'Invalid active-run identity.' }
    return $state
  } catch { $script:ActiveStateError = $_.Exception.Message; return $null }
}

function Get-MatchingAgentProcesses {
  if (-not [string]::IsNullOrWhiteSpace($ProcessSnapshotPath)) {
    $snapshot = Get-Content -LiteralPath $ProcessSnapshotPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    return $snapshot
  }

  return @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction Stop | ForEach-Object {
    [pscustomobject]@{
      ProcessId    = [int]$_.ProcessId
      Name         = [string]$_.Name
      CommandLine  = [string]$_.CommandLine
      CreationDate = $_.CreationDate
    }
  })
}

function Write-HealthEvent {
  param([int]$EventId, [string]$EntryType, [hashtable]$Fields)

  $eventsPath = Join-Path -Path $LogRoot -ChildPath 'Events'
  $jsonlPath = Join-Path -Path $eventsPath -ChildPath 'patch-agent.jsonl'
  if (-not (Test-Path -LiteralPath $eventsPath) -and $PSCmdlet.ShouldProcess($eventsPath, 'Create health event directory')) {
    New-Item -ItemType Directory -Path $eventsPath -Force | Out-Null
  }

  $record = [ordered]@{
    schemaVersion = '1.0'
    timestamp     = (Get-Date).ToString('o')
    computerName  = $env:COMPUTERNAME
    mode          = 'Machine'
    runId         = "health-$((Get-Date).ToString('yyyyMMdd-HHmmss'))"
    eventId       = $EventId
    entryType     = $EntryType
  }
  foreach ($key in $Fields.Keys) { $record[$key] = $Fields[$key] }

  if ($PSCmdlet.ShouldProcess($jsonlPath, 'Write health evidence')) {
    ($record | ConvertTo-Json -Compress -Depth 8) | Add-Content -LiteralPath $jsonlPath -Encoding UTF8
  }

  if (-not $ProcessSnapshotPath -and -not $TaskSnapshotPath -and [System.Diagnostics.EventLog]::SourceExists('DanmadePatchAgent') -and $PSCmdlet.ShouldProcess('Application/DanmadePatchAgent', 'Write health event log entry')) {
    try {
      $type = [System.Diagnostics.EventLogEntryType]::$EntryType
      Write-EventLog -LogName Application -Source 'DanmadePatchAgent' -EventId $EventId -EntryType $type -Message ($record | ConvertTo-Json -Compress -Depth 8)
    } catch { }
  }
}

function Get-AgentLogHealth {
  param([object[]]$Records, [object[]]$Processes, [string]$ActiveRunId)

  $assessed = @('Completed', 'CompletedWithFailures', 'CompletedRestartRequired', 'CompletedRestartVerificationInconclusive', 'NoPackagesSelected')
  $terminal = $assessed + @('DisabledByPolicy', 'SkippedOutsideMaintenanceWindow', 'PreflightFailed', 'UnhandledAgentError', 'InventoryInconclusive', 'PolicyLoadFailedNoPatching', 'MachineModeNotRunningAsSystem', 'LocalBootstrapFailed')
  $lastTerminal = $null; $lastAssessment = $null; $lastSkipped = $null
  $incomplete = @{}
  foreach ($record in $Records) {
    if ((Get-PropertyValue $record 'mode') -ne 'Machine') { continue }
    $time = [datetime]$record.timestamp
    if ($record.status -eq 'RunStarted') { $incomplete[[string]$record.runId] = $time }
    if ($record.status -in $terminal) {
      if (-not $lastTerminal -or $time -gt $lastTerminal) { $lastTerminal = $time }
      $incomplete.Remove([string]$record.runId)
    }
    if ($record.status -in $assessed -and (-not $lastAssessment -or $time -gt $lastAssessment)) { $lastAssessment = $time }
    if ($record.status -eq 'SkippedOutsideMaintenanceWindow' -and (-not $lastSkipped -or $time -gt $lastSkipped)) { $lastSkipped = $time }
  }
  $current = @(); $historical = @()
  foreach ($id in $incomplete.Keys) {
    $start = $incomplete[$id]
    $matching = @($Processes | Where-Object { [math]::Abs((([datetime]$_.CreationDate) - $start).TotalSeconds) -lt 30 })
    # A later completed assessment plus no corresponding live identity retires old log-only warnings.
    if ($lastAssessment -and $start -lt $lastAssessment -and $id -ne $ActiveRunId -and $matching.Count -eq 0) {
      $historical += [pscustomobject]@{ RunId = $id; StartedAt = $start }
    } else { $current += [pscustomobject]@{ RunId = $id; StartedAt = $start } }
  }
  return [pscustomobject]@{ LastTerminal = $lastTerminal; LastAssessment = $lastAssessment; LastSkipped = $lastSkipped; Incomplete = $current; HistoricalIncomplete = $historical }
}

$maxRunMinutes = Get-PolicyMaxRunMinutes
$script:ActiveStateError = $null
$activeRun = Get-ActiveRunState
$now = Get-Date
$findings = 0
$refreshTaskName = 'Danmade Patch Agent - Refresh Local Cache'
$cachePointer = Join-Path (Split-Path -Parent $PSScriptRoot) 'current.json'
$hasCachePointer = Test-Path -LiteralPath $cachePointer
if ($hasCachePointer) {
  try {
    $release = Get-Content -LiteralPath $cachePointer -Raw | ConvertFrom-Json
    $verified = [datetime]::MinValue
    if (-not [datetime]::TryParse([string](Get-PropertyValue $release 'verifiedAt' ''), [ref]$verified) -or $verified -lt $now.AddHours(-168) -or $verified -gt $now.AddMinutes(5)) { throw 'Local policy verification is missing, stale, or future-dated.' }
    if ($verified -le $now.AddHours(-120)) {
      $findings++
      Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{
        status = 'LocalPolicyVerificationAging'; verifiedAt = $verified.ToString('o')
        expiresAt = $verified.AddHours(168).ToString('o'); response = 'ReportOnlyNoTermination'
      }
    }
  } catch {
    $findings++
    Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'LocalPolicyRefreshRequired'; message = $_.Exception.Message; response = 'ReportOnlyNoTermination' }
  }
}
if ($script:ActiveStateError) {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'HealthStateUnreadable'; message = $script:ActiveStateError; response = 'ReportOnlyNoTermination' }
}
$matches = @(
  Get-MatchingAgentProcesses | Where-Object {
    $_.Name -match '^(?i)powershell\.exe$' -and $_.CommandLine -match '(?i)(danmade-patch-agent|Invoke-DanmadePatchAgentCached)\.ps1' -and $_.CommandLine -match '(?i)-Mode\s+Machine'
  }
)

foreach ($process in $matches) {
  try { $startedAt = [datetime]$process.CreationDate } catch { continue }
  $ageMinutes = [math]::Round(($now - $startedAt).TotalMinutes, 1)
  if ($ageMinutes -le $maxRunMinutes) { continue }
  $findings++

  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{
    status          = 'StaleAgentProcessDetected'
    processId       = [int]$process.ProcessId
    processStartedAt = $startedAt.ToString('o')
    ageMinutes      = $ageMinutes
    maxRunMinutes   = $maxRunMinutes
    commandLine     = [string]$process.CommandLine
    activeRunId     = if ($activeRun) { [string]$activeRun.runId } else { $null }
    activeRunPid    = if ($activeRun) { $activeRun.processId } else { $null }
    activeRunStartedAt = if ($activeRun) { [string]$activeRun.startedAt } else { $null }
    response        = 'ReportOnlyNoTermination'
  }
}

if ($activeRun -and @($matches | Where-Object { [int]$_.ProcessId -eq [int]$activeRun.processId }).Count -eq 0) {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{
    status = 'OrphanedActiveRunDetected'
    activeRunId = [string]$activeRun.runId
    processId = $activeRun.processId
    response = 'ReportOnlyNoTermination'
  }
}

try {
  $tasks = if ($TaskSnapshotPath) {
    @(Get-Content -LiteralPath $TaskSnapshotPath -Raw | ConvertFrom-Json)
  } elseif ($ProcessSnapshotPath) { @() } else {
    @(foreach ($name in @('Danmade Patch Agent - Machine', $refreshTaskName)) {
      try {
        $task = Get-ScheduledTask -TaskName $name -ErrorAction Stop
        $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop
        [pscustomobject]@{ Name = $name; State = [string]$task.State; LastRunTime = $info.LastRunTime; LastTaskResult = $info.LastTaskResult }
      } catch {
        $findings++
        Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'HealthTaskReadbackUnavailable'; taskName = $name; message = $_.Exception.Message; response = 'ReportOnlyNoTermination' }
      }
    })
  }
  foreach ($task in $tasks) {
    if ($task.Name -eq 'Danmade Patch Agent - Health') { continue }
    if ($task.Name -eq $refreshTaskName) {
      if (-not $hasCachePointer) {
        $findings++
        Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'LocalPolicyRefreshRequired'; message = 'The local cache pointer is missing.'; response = 'ReportOnlyNoTermination' }
      }
      $result = [long]$task.LastTaskResult
      # Running and not-yet-run scheduler codes are not completed launch failures.
      if ($task.State -notin @('Running', 'Queued') -and $result -notin @(0, 0x41301, 0x41303, 0x41325)) {
        $findings++
        Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{
          status = 'CacheRefreshTaskFailed'; taskName = $task.Name; taskState = $task.State
          lastRunTime = $task.LastRunTime; lastTaskResult = $result; response = 'ReportOnlyNoTermination'
        }
      }
      continue
    }
    if ($task.State -eq 'Running' -or [long]$task.LastTaskResult -ne 0) {
      $findings++
      Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{
        status = 'AgentTaskNeedsReview'
        taskName = $task.Name
        taskState = $task.State
        lastRunTime = $task.LastRunTime
        lastTaskResult = $task.LastTaskResult
        response = 'ReportOnlyNoTermination'
      }
    }
  }
} catch {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'HealthTaskReadbackUnavailable'; message = $_.Exception.Message }
}

$jsonlPath = Join-Path $LogRoot 'Events\patch-agent.jsonl'
$records = @()
if (Test-Path -LiteralPath $jsonlPath) {
  foreach ($line in (Get-Content -LiteralPath $jsonlPath -Tail 3000)) {
    try {
      $record = $line | ConvertFrom-Json -ErrorAction Stop
      $null = [datetime]$record.timestamp
      $records += $record
    } catch {
      $findings++
      Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'HealthLogRecordUnreadable' }
      break
    }
  }
}
$logHealth = Get-AgentLogHealth -Records $records -Processes $matches -ActiveRunId ([string](Get-PropertyValue $activeRun 'runId' ''))
foreach ($run in $logHealth.Incomplete) {
  if (($now - $run.StartedAt).TotalMinutes -gt $maxRunMinutes) {
    $findings++
    Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'IncompleteAgentRunDetected'; activeRunId = $run.RunId; startedAt = $run.StartedAt; response = 'ReportOnlyNoTermination' }
  }
}
if (-not $logHealth.LastTerminal -or ($now - $logHealth.LastTerminal).TotalHours -gt 26) {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'AgentRunEvidenceMissing'; lastTerminalAt = $logHealth.LastTerminal; thresholdHours = 26; response = 'ReportOnlyNoTermination' }
}
if (-not $logHealth.LastAssessment -or ($now - $logHealth.LastAssessment).TotalHours -gt 26) {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'AgentAssessmentOverdue'; lastAssessmentAt = $logHealth.LastAssessment; lastSkippedAt = $logHealth.LastSkipped; thresholdHours = 26; response = 'ReportOnlyNoTermination' }
}
Write-HealthEvent -EventId 5002 -EntryType Information -Fields @{
  status = 'HealthCheckCompleted'
  findings = $findings
  matchingProcessCount = $matches.Count
  maxRunMinutes = $maxRunMinutes
  lastAgentTerminalAt = $logHealth.LastTerminal
  lastAssessmentAt = $logHealth.LastAssessment
  lastSkippedAt = $logHealth.LastSkipped
  historicalIncompleteCount = @($logHealth.HistoricalIncomplete).Count
  response = 'ReportOnlyNoTermination'
}
