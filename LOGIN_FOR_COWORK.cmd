@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo OMH M365 Agent - Cowork login helper
echo.
echo This helper creates the Microsoft login cache in this linux-llm folder.
echo Run this yourself. Do not ask an LLM to read, copy, or repeat the device code.
echo.

if not exist ".env" (
  echo Missing .env in this folder.
  echo Create .env from .env.example with approved Azure Application values, then run this helper again.
  echo Do not paste client ID or tenant ID into chat.
  pause
  exit /b 3
)

set "USE_NODE="
set "EXE=%~dp0omh-m365.exe"
if exist "%EXE%" goto run_login

set "EXE=%~dp0..\win-x64\omh-m365.exe"
if exist "%EXE%" goto run_login

where node >nul 2>nul
if "%ERRORLEVEL%"=="0" (
  if exist "%~dp0omh-m365.cjs" (
    set "USE_NODE=1"
    goto run_login
  )
)

echo Could not find a Windows login runner.
echo Expected one of:
echo   %~dp0omh-m365.exe
echo   %~dp0..\win-x64\omh-m365.exe
echo   node.exe on PATH plus %~dp0omh-m365.cjs
echo.
echo Ask the distributor to include the Windows helper executable or install Node.js.
pause
exit /b 2

:run_login
if /I "%~1"=="--check" (
  echo Login helper check passed.
  if "%USE_NODE%"=="1" (
    echo Runner: node %~dp0omh-m365.cjs
  ) else (
    echo Runner: %EXE%
  )
  echo Cache target folder: %~dp0.cache
  exit /b 0
)

echo Opening Microsoft device login page...
start "" "https://microsoft.com/devicelogin"
echo.
echo A Microsoft device code will appear below.
echo Enter it yourself in the browser. Do not paste the code into chat.
echo.

if "%USE_NODE%"=="1" (
  node "%~dp0omh-m365.cjs" auth login
) else (
  "%EXE%" auth login
)

if errorlevel 1 (
  echo.
  echo Login failed or was cancelled.
  pause
  exit /b 1
)

echo.
echo Checking login status...
if "%USE_NODE%"=="1" (
  node "%~dp0omh-m365.cjs" auth status
) else (
  "%EXE%" auth status
)

echo.
echo If loggedIn is true, return to Cowork and ask it to continue.
pause
endlocal
