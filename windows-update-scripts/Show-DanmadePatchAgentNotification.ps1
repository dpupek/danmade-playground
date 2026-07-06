<#
.SYNOPSIS
Shows logged-on user notifications for Danmade Patch Agent restart or failure state.

.DESCRIPTION
This script is intended to run from a user-context Group Policy scheduled task.
It reads pending notification state written by danmade-patch-agent.ps1, shows a
bounded Windows popup, and records per-user acknowledgement metadata so the same
run does not nag on every logon or unlock.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$MachineStatePath = 'C:\ProgramData\DanmadePatchAgent\State\pending-notification.json',

  [string]$UserStatePath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'DanmadePatchAgent\State\pending-notification.json'),

  [string]$AckRoot = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'DanmadePatchAgent\Notifications'),

  [int]$RepeatHours = 4,

  [int]$MaxShowCount = 12,

  [int]$PopupTimeoutSeconds = 90,

  [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Directory {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return }
  if (Test-Path -LiteralPath $Path) { return }
  if ($PSCmdlet.ShouldProcess($Path, 'Create directory')) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function ConvertTo-SafeFileName {
  param([string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) { return 'unknown' }
  return ($Value -replace '[^A-Za-z0-9._-]', '_')
}

function Read-NotificationState {
  param(
    [string]$Path,
    [string]$Scope
  )

  if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
  if (-not (Test-Path -LiteralPath $Path)) { return $null }

  try {
    $state = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if (-not $state.PSObject.Properties['status']) { return $null }
    if (-not $state.PSObject.Properties['runId']) { return $null }
    $state | Add-Member -NotePropertyName StatePath -NotePropertyValue $Path -Force
    $state | Add-Member -NotePropertyName StateScope -NotePropertyValue $Scope -Force
    return $state
  } catch {
    Write-Warning "Unable to read Danmade Patch Agent notification state from '$Path': $($_.Exception.Message)"
    return $null
  }
}

function Get-StateId {
  param([object]$State)

  $parts = @(
    [string]$State.computerName,
    [string]$State.mode,
    [string]$State.runId,
    [string]$State.status,
    [string]$State.StateScope
  )
  return (ConvertTo-SafeFileName -Value ($parts -join '-'))
}

function Get-AckPath {
  param([object]$State)
  return (Join-Path -Path $AckRoot -ChildPath ("{0}.json" -f (Get-StateId -State $State)))
}

function Read-Ack {
  param([object]$State)

  $path = Get-AckPath -State $State
  if (-not (Test-Path -LiteralPath $path)) {
    return [pscustomobject]@{
      path      = $path
      showCount = 0
      lastShown = $null
    }
  }

  try {
    $ack = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $count = 0
    if ($ack.PSObject.Properties['showCount']) { $count = [int]$ack.showCount }
    $lastShown = $null
    if ($ack.PSObject.Properties['lastShown'] -and -not [string]::IsNullOrWhiteSpace([string]$ack.lastShown)) {
      $lastShown = [datetime]::Parse([string]$ack.lastShown)
    }
    return [pscustomobject]@{
      path      = $path
      showCount = $count
      lastShown = $lastShown
    }
  } catch {
    return [pscustomobject]@{
      path      = $path
      showCount = 0
      lastShown = $null
    }
  }
}

function Test-RestartAlreadySatisfied {
  param([object]$State)

  if ([string]$State.status -ne 'CompletedRestartRequired') { return $false }
  if (-not $State.PSObject.Properties['timestamp']) { return $false }

  try {
    $stateTime = [datetime]::Parse([string]$State.timestamp)
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $lastBoot = if ($os.LastBootUpTime -is [datetime]) {
      [datetime]$os.LastBootUpTime
    } else {
      [Management.ManagementDateTimeConverter]::ToDateTime([string]$os.LastBootUpTime)
    }
    return ($lastBoot -gt $stateTime)
  } catch {
    return $false
  }
}

function Test-ShouldShow {
  param(
    [object]$State,
    [object]$Ack
  )

  if ($Force) { return $true }
  if (Test-RestartAlreadySatisfied -State $State) { return $false }
  if ($Ack.showCount -ge $MaxShowCount) { return $false }
  if ($null -eq $Ack.lastShown) { return $true }

  $nextShow = $Ack.lastShown.AddHours($RepeatHours)
  return ((Get-Date) -ge $nextShow)
}

function Get-PackageSummaryText {
  param([object[]]$Packages)

  $ids = @(
    foreach ($pkg in @($Packages)) {
      if ($null -eq $pkg) { continue }
      if ($pkg.PSObject.Properties['id'] -and -not [string]::IsNullOrWhiteSpace([string]$pkg.id)) {
        [string]$pkg.id
      } elseif ($pkg.PSObject.Properties['name'] -and -not [string]::IsNullOrWhiteSpace([string]$pkg.name)) {
        [string]$pkg.name
      }
    }
  )

  if ($ids.Count -eq 0) { return '' }
  if ($ids.Count -le 5) { return ($ids -join ', ') }
  return (($ids | Select-Object -First 5) -join ', ') + " and $($ids.Count - 5) more"
}

function New-NotificationMessage {
  param([object]$State)

  $modeText = if ($State.PSObject.Properties['mode']) { [string]$State.mode } else { 'Patch' }
  if ([string]$State.status -eq 'CompletedRestartRequired') {
    $packages = Get-PackageSummaryText -Packages @($State.restartRequiredPackages)
    $detail = if ($packages) { "`r`n`r`nUpdated package(s): $packages" } else { '' }
    return [pscustomobject]@{
      Title   = 'Restart required for application updates'
      Message = "Application updates were installed on $env:COMPUTERNAME and Windows needs a restart to finish applying them.$detail`r`n`r`nPlease restart today."
      Icon    = 48
      Reason  = "$modeText restart required"
    }
  }

  $failedPackages = @()
  if ($State.PSObject.Properties['failedPackages']) { $failedPackages = @($State.failedPackages) }
  $packagesText = Get-PackageSummaryText -Packages $failedPackages
  $detailText = if ($packagesText) { "`r`n`r`nPackage(s): $packagesText" } else { '' }
  return [pscustomobject]@{
    Title   = 'Application update issue'
    Message = "Some application updates did not complete on $env:COMPUTERNAME. You can keep working; Nexport IT has details in the patch-agent logs.$detailText"
    Icon    = 48
    Reason  = "$modeText update issue"
  }
}

function Show-UserPopup {
  param([object]$Notification)

  if (-not $PSCmdlet.ShouldProcess($Notification.Title, 'Show Danmade Patch Agent notification')) {
    Write-Output $Notification
    return
  }

  $shell = New-Object -ComObject WScript.Shell
  $buttonType = 0
  [void]$shell.Popup([string]$Notification.Message, $PopupTimeoutSeconds, [string]$Notification.Title, ($buttonType + [int]$Notification.Icon))
}

function Write-Ack {
  param(
    [object]$State,
    [object]$Ack,
    [object]$Notification
  )

  Ensure-Directory -Path $AckRoot
  $nextCount = [int]$Ack.showCount + 1
  $record = [ordered]@{
    schemaVersion = '1.0'
    computerName  = [string]$State.computerName
    mode          = [string]$State.mode
    runId         = [string]$State.runId
    status        = [string]$State.status
    stateScope    = [string]$State.StateScope
    statePath     = [string]$State.StatePath
    showCount     = $nextCount
    lastShown     = (Get-Date).ToString('o')
    userName      = "$env:USERDOMAIN\$env:USERNAME"
    reason        = [string]$Notification.Reason
  }

  if ($PSCmdlet.ShouldProcess($Ack.path, 'Write notification acknowledgement')) {
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Ack.path -Encoding utf8
  }

  $logPath = Join-Path -Path $AckRoot -ChildPath 'notification.jsonl'
  if ($PSCmdlet.ShouldProcess($logPath, 'Append notification log')) {
    ($record | ConvertTo-Json -Compress -Depth 8) | Add-Content -LiteralPath $logPath -Encoding utf8
  }
}

$states = @(
  Read-NotificationState -Path $MachineStatePath -Scope 'Machine'
  Read-NotificationState -Path $UserStatePath -Scope 'User'
) | Where-Object { $null -ne $_ }

foreach ($state in $states) {
  if ([string]$state.status -notin @('CompletedRestartRequired', 'CompletedWithFailures', 'PreflightFailed', 'MachineModeNotRunningAsSystem', 'UnhandledAgentError')) {
    continue
  }

  $ack = Read-Ack -State $state
  if (-not (Test-ShouldShow -State $state -Ack $ack)) { continue }

  $notification = New-NotificationMessage -State $state
  Show-UserPopup -Notification $notification
  Write-Ack -State $state -Ack $ack -Notification $notification
}

# SIG # Begin signature block
# MIIgEQYJKoZIhvcNAQcCoIIgAjCCH/4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdnsBOJcbIPI+3
# pU+2EEDYZ6c6PQLlVOecly54rxdv7qCCGiwwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJWeRn6flWOE
# 2s22fsyELeXabbrvzw3RLcS5u/ndlXIrMA0GCSqGSIb3DQEBAQUABIIBACFDRpFJ
# VhMIWFMQ9UsLoxwWqvXP+CN2zkatlAcoXk/Zc3s85rMu0kLLxQDQjLzW9FGvH/gS
# n01EpiHC+d+5ecNB+KkXcHyB/Ju/usaGtwxo814mRC/74iTLRePx8aPjvhYOqa1M
# dqPOK2jZGoAs4UlfPay2a2WygbXvvYCgr6BVgGAjymYvKC7H5i0kioH+yXGiSu/j
# J28pKyjOzJln8wp8vYbNN96eAv+JKq80jGmBvxzWRCwjEBFdehodOcVHUJhYuP5T
# mbmvhEKHi+4HmfiAH6EV+hqOfllYFBWje1CFb83u+AX/uOttQqd14a0kEy2KZZ5R
# naX3+0AD5DAuX76hggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExAhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNzA2MTYyNjU0
# WjAvBgkqhkiG9w0BCQQxIgQgvh55oHZ73IgWb2/mqfiG+RFMgd1BNkqJTWjfHDbP
# ZuswDQYJKoZIhvcNAQEBBQAEggIAq794OJrKuPTeFPj6yYsDG6DdaHWicvK8MshU
# o12USMGHECLIFCwFD98LEf8RcchxV7EBQdr3aOHsX3OmP6Olot21lt1BmGoQHYdp
# gSUKlOQxsEZ17PlKtPTXUNpxi8BTDcelpFAvboYwy3c/FF12Xi7/6zaa4JCGUqtS
# 2JgQdbr34P0GFk++XlHlxugdjdkvX1F4O7rk5fuOLAILYHHf4pTYlFlkSdP/rZ8a
# 6PPSu0LDCd+ZmPoOKv8hRvw5Si3XQvOif95oC07XtiWsBB4Zo/PRp/SzXziOzfNA
# CvKXVX2nrXIsPgjNhes0l+gBxXjBU++TRBRb8X9LvwE8bp9GlYCQDQsFNKitpkBW
# XvSA9Su4+6iZVzB+/5DtO1/os10oJEDgtyzy5fYgHCjMbkoHx2x3vNnbKujQS3lZ
# iOigM+VNLdH6UYdJdeaRWe/H+KSKxF2WOHPdVfs8wTyeWcjPDYu58CnkkpiswA1k
# AodGXVnT6/TBnbHSGernjQSH7645agxxeQ3zS5aAlxVAjPB7XfgKPB8AYSQhfEXf
# wQBUsY/Tw4/7PqiSRKLFllFKlDSPnxKar4nPXZ40PMoF2u3+SRJKGa4iAtQZ7lsH
# GZnYnF4qR3N1onilAWzcbLLsdNDGgnIlTTc5mHgH4Hu0JHXsoTMkA1f3MeMVLpvH
# PSz7vrU=
# SIG # End signature block
