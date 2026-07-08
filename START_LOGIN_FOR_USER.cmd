@echo off
setlocal
cd /d "%~dp0"
echo OMH M365 Agent - Microsoft login
echo.
echo A Microsoft device login message will appear in this window.
echo The human user must enter the displayed code in the browser.
echo LLMs must not copy, repeat, store, or summarize the device code.
echo.
start "" "https://microsoft.com/devicelogin"
"%~dp0omh-m365.exe" auth login
endlocal
