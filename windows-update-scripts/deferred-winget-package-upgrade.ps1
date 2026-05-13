param(
  [Parameter(Mandatory = $true)]
  [string[]]$PackageId
)

$ErrorActionPreference = 'Stop'

$WingetCommonArgs = @(
  '--include-unknown',
  '--silent',
  '--exact',
  '--accept-package-agreements',
  '--accept-source-agreements',
  '--source','winget'
)

function Wait-ForPwshExit {
  $lastSignature = $null

  while ($true) {
    $pwshProcesses = @(Get-Process pwsh -ErrorAction SilentlyContinue | Sort-Object Id)
    if ($pwshProcesses.Count -eq 0) { return }

    $signature = ($pwshProcesses | ForEach-Object { "{0}:{1}" -f $_.Id, $_.ProcessName }) -join ','
    if ($signature -ne $lastSignature) {
      Write-Host ''
      Write-Host 'Waiting for PowerShell 7 processes to exit before running deferred upgrades...' -ForegroundColor Yellow
      Write-Host 'Close the windows below and check Task Manager for lingering pwsh.exe processes if needed.' -ForegroundColor Yellow
      foreach ($proc in $pwshProcesses) {
        $windowTitle = if ([string]::IsNullOrWhiteSpace($proc.MainWindowTitle)) { '(no window title)' } else { $proc.MainWindowTitle.Trim() }
        $path = $null
        try { $path = $proc.Path } catch { $path = $null }
        if ($path) {
          Write-Host ("- PID {0}: {1} | {2}" -f $proc.Id, $windowTitle, $path)
        } else {
          Write-Host ("- PID {0}: {1}" -f $proc.Id, $windowTitle)
        }
      }
      $lastSignature = $signature
    }

    Start-Sleep -Seconds 2
  }
}

try {
  $winget = Get-Command winget -ErrorAction SilentlyContinue
  if (-not $winget) {
    Write-Host 'winget was not found on PATH. Cannot run deferred upgrades.' -ForegroundColor Red
    exit 1
  }

  Wait-ForPwshExit

  $logDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winget-update-script-logs'
  if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
  }

  $failures = @()
  foreach ($id in @($PackageId | Select-Object -Unique)) {
    $safeId = ($id -replace '[^A-Za-z0-9._-]','_')
    $logPath = Join-Path -Path $logDir -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $args = @('upgrade','--id',$id,'--log',$logPath) + $WingetCommonArgs

    Write-Host ''
    Write-Host ("Running deferred upgrade: winget {0}" -f ($args -join ' ')) -ForegroundColor Cyan
    & $winget.Path @args
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
      $failures += [pscustomobject]@{
        Id = $id
        ExitCode = $exitCode
        LogPath = $logPath
      }
      Write-Host ("Deferred upgrade failed for {0}. Exit code: {1}" -f $id, $exitCode) -ForegroundColor Red
      Write-Host ("Log: {0}" -f $logPath) -ForegroundColor Yellow
    } else {
      Write-Host ("Deferred upgrade completed successfully for {0}." -f $id) -ForegroundColor Green
    }
  }

  if ($failures.Count -gt 0) {
    Write-Host ''
    Write-Host 'Deferred upgrade summary:' -ForegroundColor Yellow
    foreach ($failure in $failures) {
      Write-Host ("- {0}: exit code {1}. Log: {2}" -f $failure.Id, $failure.ExitCode, $failure.LogPath)
    }
    exit 1
  }

  Write-Host ''
  Write-Host 'All deferred upgrades completed successfully.' -ForegroundColor Green
  exit 0
} catch {
  Write-Host ("Deferred upgrade helper failed unexpectedly: {0}" -f $_.Exception.Message) -ForegroundColor Red
  exit 1
} finally {
  Read-Host 'Press Enter to close this window' | Out-Null
}
