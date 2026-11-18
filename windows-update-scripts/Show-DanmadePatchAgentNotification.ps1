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

function Read-ApplicationRestartNotifications {
  $path = Join-Path (Split-Path -Parent $UserStatePath) 'pending-application-restart.json'
  if (-not (Test-Path -LiteralPath $path)) { return }
  try {
    $pending = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    foreach ($package in @($pending.packages)) {
      if ([datetimeoffset]$package.expiresAt -le [datetimeoffset]::Now) { continue }
      [pscustomobject]@{
        computerName = $env:COMPUTERNAME; mode = 'User'; StateScope = 'User'; StatePath = $path
        runId = "$($package.runId)-$($package.id)-$($package.source)-$($package.available)"
        status = 'ApplicationRestartPendingVerification'; applicationPackage = $package
      }
    }
  } catch { Write-Warning "Unable to read application restart state: $($_.Exception.Message)" }
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
  if ($State.status -eq 'ApplicationRestartPendingVerification') {
    $app = if ($State.applicationPackage.name) { $State.applicationPackage.name } else { $State.applicationPackage.id }
    return [pscustomobject]@{
      Title = 'Reopen an application to finish updating'
      Message = "Save your work in $app, then close and reopen it when convenient. No Windows restart is needed for this update."
      Icon = 48
      Reason = 'Application restart pending verification'
    }
  }
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
  Read-ApplicationRestartNotifications
) | Where-Object { $null -ne $_ }

foreach ($state in $states) {
  if ([string]$state.status -notin @('CompletedRestartRequired', 'CompletedWithFailures', 'PreflightFailed', 'MachineModeNotRunningAsSystem', 'UnhandledAgentError', 'ApplicationRestartPendingVerification')) {
    continue
  }

  $ack = Read-Ack -State $state
  if (-not (Test-ShouldShow -State $state -Ack $ack)) { continue }

  $notification = New-NotificationMessage -State $state
  Show-UserPopup -Notification $notification
  Write-Ack -State $state -Ack $ack -Notification $notification
}
