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
    return @(Get-Content -LiteralPath $ProcessSnapshotPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
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

  if ([System.Diagnostics.EventLog]::SourceExists('DanmadePatchAgent') -and $PSCmdlet.ShouldProcess('Application/DanmadePatchAgent', 'Write health event log entry')) {
    try {
      $type = [System.Diagnostics.EventLogEntryType]::$EntryType
      Write-EventLog -LogName Application -Source 'DanmadePatchAgent' -EventId $EventId -EntryType $type -Message ($record | ConvertTo-Json -Compress -Depth 8)
    } catch { }
  }
}

$maxRunMinutes = Get-PolicyMaxRunMinutes
$script:ActiveStateError = $null
$activeRun = Get-ActiveRunState
$now = Get-Date
$findings = 0
$cachePointer = Join-Path (Split-Path -Parent $PSScriptRoot) 'current.json'
if (Test-Path -LiteralPath $cachePointer) {
  try {
    $release = Get-Content -LiteralPath $cachePointer -Raw | ConvertFrom-Json
    $verified = [datetime]::MinValue
    if (-not [datetime]::TryParse([string](Get-PropertyValue $release 'verifiedAt' ''), [ref]$verified) -or $verified -lt $now.AddHours(-168) -or $verified -gt $now.AddMinutes(5)) { throw 'Local policy verification is missing, stale, or future-dated.' }
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
    @(foreach ($name in @('Danmade Patch Agent - Machine', 'Danmade Patch Agent - Health')) {
      $task = Get-ScheduledTask -TaskName $name -ErrorAction Stop
      $info = $task | Get-ScheduledTaskInfo -ErrorAction Stop
      [pscustomobject]@{ Name = $name; State = [string]$task.State; LastRunTime = $info.LastRunTime; LastTaskResult = $info.LastTaskResult }
    })
  }
  foreach ($task in $tasks) {
    if ($task.Name -eq 'Danmade Patch Agent - Health') { continue }
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
$terminalStatuses = @('Completed', 'CompletedWithFailures', 'CompletedRestartRequired', 'NoPackagesSelected', 'DisabledByPolicy', 'SkippedOutsideMaintenanceWindow', 'PreflightFailed', 'UnhandledAgentError', 'InventoryInconclusive', 'PolicyLoadFailedNoPatching', 'MachineModeNotRunningAsSystem', 'LocalBootstrapFailed')
$lastTerminal = $null
$incomplete = @{}
if (Test-Path -LiteralPath $jsonlPath) {
  foreach ($line in (Get-Content -LiteralPath $jsonlPath -Tail 3000)) {
    try {
      $record = $line | ConvertFrom-Json -ErrorAction Stop
      if ($record.mode -eq 'Machine' -and $record.status -eq 'RunStarted') { $incomplete[[string]$record.runId] = [datetime]$record.timestamp }
      if ($record.mode -eq 'Machine' -and $record.status -in $terminalStatuses) {
        $lastTerminal = [datetime]$record.timestamp
        $incomplete.Remove([string]$record.runId)
      }
    } catch {
      $findings++
      Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'HealthLogRecordUnreadable' }
      break
    }
  }
}
foreach ($id in $incomplete.Keys) {
  if (($now - $incomplete[$id]).TotalMinutes -gt $maxRunMinutes) {
    $findings++
    Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'IncompleteAgentRunDetected'; activeRunId = $id; startedAt = $incomplete[$id]; response = 'ReportOnlyNoTermination' }
  }
}
if (-not $lastTerminal -or ($now - $lastTerminal).TotalHours -gt 26) {
  $findings++
  Write-HealthEvent -EventId 5601 -EntryType Warning -Fields @{ status = 'AgentRunEvidenceMissing'; lastTerminalAt = $lastTerminal; thresholdHours = 26; response = 'ReportOnlyNoTermination' }
}
Write-HealthEvent -EventId 5002 -EntryType Information -Fields @{
  status = 'HealthCheckCompleted'
  findings = $findings
  matchingProcessCount = $matches.Count
  maxRunMinutes = $maxRunMinutes
  lastAgentTerminalAt = $lastTerminal
  response = 'ReportOnlyNoTermination'
}
# SIG # Begin signature block
# MIIgEQYJKoZIhvcNAQcCoIIgAjCCH/4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsXhz7vZNA6zXo
# Nl0ySCT0d1YKv6NY01ENCoTRdUOnuqCCGiwwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# ubdcMIIG7TCCBNWgAwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsF
# ADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNV
# BAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hB
# MjU2IDIwMjUgQ0ExMB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJE
# aWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYg
# MTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz
# /TjGMPh8CQ6RucZCLv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMx
# t9l3Dm5oCG5Z69AqEN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkU
# xDjLJhBI+qiZP3aixAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZE
# eoy/0B8Zm+nzKBgSWG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEO
# vZ3fP2VxIgeR28Y90ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn
# /FnbuTegcoRcFEZ9FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolv
# ykE1jive4PuESaT8zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZ
# CZVl6pNUOCybMc/kxC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW8
# 3vV8d67WowuUs0HdY4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjd
# zi0xlp70W+ow63/0++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9ic
# flf55W1zOEKAcJVAIXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0G
# A1UdDgQWBBQUyWOKMC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEy
# NTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hB
# MjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQI
# ENoyKqxrFbJblzCVP/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGh
# z45F4Wt4WGdNdIbRWt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfH
# hHv1mx9EVh71SJelcEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA
# 02nXD1nYLI7Ijmiv+xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7K
# VyF5aQyWP4KhVGUUTnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jht
# KVnd3D0rbHXEu27o+Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvK
# qzpAec9etE14jOQAzLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs
# 0NzU/+xX7NbEdcofy0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGR
# sMNbXmMREUTgu+Oca7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJ
# Nt68Yfim6cjEzVSiLSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yW
# QPT9gzCCBu4wggXWoAMCAQICE1YAAAD7uI/G5fXiwIQAAQAAAPswDQYJKoZIhvcN
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKKPekP3cgMP
# PoPR+c9KwmwHO8cVGlb5ayXqq0kK+syXMA0GCSqGSIb3DQEBAQUABIIBAJkjEEec
# 6+G+joS6y2dxud3lQh91zwD6qV5WWNJr5Q7xT2nYnOc6TY82qPECfQxGAeeF2lRc
# /iB9xRlteUNpvsuw8mBCZGpsXb1E31WISHwhllSuRLitlqPxP691Kh3AUBDTwAHX
# lFW/rJKCj2vxut4VLMGJXEiHE+RABVRq3w/z53ysGvc/zxnil+/875LKkN2Rg/fZ
# 9WDYPbA2SRDk0mOkSY2CKt8a0pkD15SzCN21IhV9eu65L4ebSZ0QI6wKGktYldEg
# EGQZo/4eqWu9IrxmpR0nG3RUSkY76uWIJleIUs4EB2UsVNFAxfOdRNS4V2dSBPjO
# uZk8pAGh72exD5KhggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExAhAIT9wzT35FTtvDD4/5khg1MA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYxMDE2MTczOTQ0
# WjAvBgkqhkiG9w0BCQQxIgQgnEcRRvOCrvQaa3/0GK3WuCNpTYRpFau5/4lqFo3V
# QaswDQYJKoZIhvcNAQEBBQAEggIASg5Lcgc5Yoqrm7CQBwjY+TerpBpUUEjfK4wh
# 5lXvBAa6Oz7XQxs5hXfPF4GwkE1AShQ4Ww/8k7CRzCWeN8kG3fLjZ/ZMFP0IceFB
# alYPrNweEs2SFIJ0iDhFkUTX3mkzNNtWYNJ1qa3Kq7AdvodNM2BTPXDBnYaqIVIr
# S8/k4hiRAVvP7n4i0NLYrCKcC8Yu9U7xbLf78OxBbYQcGGcTYauKphjsIAbOYTj5
# aIsMi9s7/ABFHaQ4ZCufHrFr5gNnsNG41E9IylRUw0cEMJQbJQN5B06k5kPUg1FY
# UEZ783nHjBEb/F3nMsAke1+ClIQzzcimoICpZ55DmjN94InFVmANkitHUo3B3lbN
# PlkNZvLVRS2QyUAb9j92f8VNTgbyY+oFYRvMN0JsIYrRx6N6UDoChpqPi9A4SMYY
# XON3Uqsujlpi25X/yfVZt1RnfChPfrTgkV2UD8hf4FdwuodPYlB2gSnKxrZZ2Sgr
# XJO8sdDAlTEZOJhUdyiGBjqorgxARBSELShOYdFX6+RsE3QijF2GCbfpv9N1gzNh
# C7LL6hUiAIgvB5gnKNgr942i3VdsMM683WbWGNVHEluUm2YXg7tIL6Nh/TZaotZy
# 8GgDxoUV95j6Jsn/OTLIHS/OVZQp1zhuX8hoDq5s4X7obi40cOCoL5XEYLDVQO1X
# U+sW1oc=
# SIG # End signature block
