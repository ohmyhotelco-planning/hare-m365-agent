@echo off
setlocal
cd /d "%~dp0"
set "OMH_M365_DIR=%CD%"
echo OMH M365 Agent - LLM startup check
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference = 'Stop';" ^
  "$exe = Join-Path $env:OMH_M365_DIR 'omh-m365.exe';" ^
  "Write-Host '[1/2] doctor';" ^
  "$doctorJson = & $exe doctor;" ^
  "if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }" ^
  "$doctorJson | Write-Output;" ^
  "$status = $doctorJson | ConvertFrom-Json;" ^
  "if (-not $status.configured) {" ^
  "  Write-Host '';" ^
  "  Write-Host 'Configuration is missing.';" ^
  "  Write-Host 'The human user must create .env from .env.example and fill approved Azure Application values locally.';" ^
  "  Write-Host 'Do not ask the user to paste client ID or tenant ID into chat.';" ^
  "  exit 3;" ^
  "}" ^
  "if (-not $status.loggedIn) {" ^
  "  Write-Host '';" ^
  "  Write-Host 'Authentication is missing or expired.';" ^
  "  Write-Host 'Before any Outlook, Teams, or Files command, complete login first:';" ^
  "  Write-Host '  START_LOGIN_FOR_USER.cmd';" ^
  "  Write-Host 'After the human says login is complete, rerun this startup check and then resume the original task.';" ^
  "  exit 4;" ^
  "}" ^
  "Write-Host '';" ^
  "Write-Host '[2/2] auth status';" ^
  "& $exe auth status;" ^
  "if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }" ^
  "Write-Host '';" ^
  "Write-Host 'Ready for read-only Microsoft 365 commands.'"
exit /b %ERRORLEVEL%
endlocal
