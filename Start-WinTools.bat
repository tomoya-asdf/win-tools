@echo off
setlocal
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File "%~dp0Start-WinTools.ps1" %*
set "RESULT=%errorlevel%"
if not "%RESULT%"=="0" pause
exit /b %RESULT%
