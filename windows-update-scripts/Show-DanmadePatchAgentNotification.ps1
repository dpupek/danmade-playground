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
  if ([string]$state.status -notin @('CompletedRestartRequired', 'CompletedWithFailures', 'PreflightFailed', 'UnhandledAgentError')) {
    continue
  }

  $ack = Read-Ack -State $state
  if (-not (Test-ShouldShow -State $state -Ack $ack)) { continue }

  $notification = New-NotificationMessage -State $state
  Show-UserPopup -Notification $notification
  Write-Ack -State $state -Ack $ack -Notification $notification
}

# SIG # Begin signature block
# MIIgCwYJKoZIhvcNAQcCoIIf/DCCH/gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAX/NqN0n46SiU3
# JzL0qTDZ+VNR13mN2/Ow1x88gzz+s6CCGiYwggWNMIIEdaADAgECAhAOmxiO+dAt
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
# CzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIPmBDnBg0oyoSEMuFQhO
# yBxVgtzmQfZABzrFM60iHgKoMA0GCSqGSIb3DQEBAQUABIIBADqftbpuj8ABIbxY
# vyHcRq8AWxV6GkNNSYmZyRL2VY9MNNFW8KY3kZ2a8ZpqJi+kpp7uxqtBkGXRCVWr
# SMdcwH/MzE4ts5mhY+Y8NxWrRmlx/hEZfcl4fMkVxK1P4iUq7p+Y/58dCEy5sRlj
# onGJ3VZn4f62osZViNC5Xq+A3ZBDen2DlfbX8+yNa3wybE2dSxr3rAdJuToXVnDb
# ZRiEeIanXLi7cIVJwTnyhFFHS1M1Ut5eqwVLCl1DKEkubJwG8Q+nBHOuztjhBeHF
# bk/y3QqdgQ1HJrvQaCbnf731snur+THoBnliQrWPUAXUTEZ0gWiMtev/r929p6lI
# kpyrQxChggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# AhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkD
# MQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNjI5MTQ0ODQ1WjAvBgkq
# hkiG9w0BCQQxIgQg0Ejy4tT0D1V4m6XEdTm11LZdSFotWlOuO+LWjAxfkycwDQYJ
# KoZIhvcNAQEBBQAEggIASy2iWFPETwrKqZkrtoef4ShvIUfKtjlkTsudbrhL1jFe
# Tsufj5P/gRJZg6RBCx6RYj0UjhyxxjS+FP7l7tnnZcSR2yLy6ytOkBIofy2qFjNF
# 5PHIXmtXUomWFOxwb3IYc0AgV1f86lSWc7adlsiUS24OEqpAbJxlvLhKZbUs0hHF
# bMzxtUWMN/D5vkICwZkSNARLG3fmnpMNyp1XV0/dpfudGgLhI5cg+awDAaFmU9lQ
# osAbNR3Jx0exwVDwOR66Jnp+R7UWHgvYENx+Vg3jySW+2YaxueUUk/y7S5kUjUhH
# 8D70Zz1k4PufuPKHeyImqUK6eQTbRlMC4O4PJ8aMZ4bVW8YId/+i6UIkcyrBXoEI
# lnJ7/fCx5HJTsY4Va7MKYBFxWNTDaRxw8fcnioIEUqJPBctxb7WjSo0yw2lmSkh6
# zTKIUSF/ZUpfJsxQykNnOvwwnFWcqauIZoGjKi2LSVjqlnNKe6QBtxbJNAkudAOD
# 0v8BQJw7+NMafdTq+rzsOA/is9lF8TU5T/3OXJ99GQdlZ8AmNPuxtZ81MCc8PXoN
# QJ0Asp8+2JsONMjT2fHOGUcrWUyZr7+NqXYK0lhKiivdZkbQSBMDQIgIEowS9dar
# Zw+3dVLjf/XA12M3xidl6jJjQQOSUiasOLDP1kxJxSBPGmXp3IuuYSy4YiA8QCs=
# SIG # End signature block
