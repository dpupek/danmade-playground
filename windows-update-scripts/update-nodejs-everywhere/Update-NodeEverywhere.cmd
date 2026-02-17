@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Update-NodeEverywhere.ps1"

if not exist "%PS1%" (
  echo ERROR: Could not find "%PS1%"
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
exit /b %ERRORLEVEL%
