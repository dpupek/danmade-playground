@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "TARGET_SCRIPT=%SCRIPT_DIR%winget-ms-store-update-all.ps1"

if not exist "%TARGET_SCRIPT%" (
  echo Target script not found:
  echo   %TARGET_SCRIPT%
  exit /b 1
)

call :find_pwsh
if defined PWSH_EXE goto run_script

echo PowerShell 7 or later was not found.
choice /C YN /N /M "Install the latest PowerShell with winget now? [Y/N]: "
if errorlevel 2 (
  echo Skipping install. Cannot continue without PowerShell 7+.
  exit /b 1
)

where winget >nul 2>nul
if errorlevel 1 (
  echo winget was not found on PATH, so PowerShell cannot be installed automatically.
  exit /b 1
)

echo Installing Microsoft.PowerShell via winget...
winget install --id Microsoft.PowerShell --exact --source winget --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
  echo PowerShell installation failed.
  exit /b 1
)

call :find_pwsh
if not defined PWSH_EXE (
  echo PowerShell appears to be installed, but pwsh.exe was not found on PATH.
  echo Open a new terminal and run this launcher again.
  exit /b 1
)

:run_script
"%PWSH_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%TARGET_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
exit /b %EXIT_CODE%

:find_pwsh
set "PWSH_EXE="

for %%I in (pwsh.exe) do set "PWSH_EXE=%%~$PATH:I"
if defined PWSH_EXE goto :eof

if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" (
  set "PWSH_EXE=%ProgramFiles%\PowerShell\7\pwsh.exe"
  goto :eof
)

if exist "%ProgramW6432%\PowerShell\7\pwsh.exe" (
  set "PWSH_EXE=%ProgramW6432%\PowerShell\7\pwsh.exe"
  goto :eof
)

if exist "%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe" (
  set "PWSH_EXE=%LocalAppData%\Microsoft\PowerShell\7\pwsh.exe"
  goto :eof
)

goto :eof
