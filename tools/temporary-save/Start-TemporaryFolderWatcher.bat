@echo off
setlocal
powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Start-TemporaryFolderWatcher.ps1" %*
exit /b %errorlevel%
