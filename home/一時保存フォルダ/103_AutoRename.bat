@echo off
setlocal

REM 自動リネーム監視スクリプトを非表示で起動
set "TARGET=%~dp0\400_RenameWatcher.ps1"

powershell -NoProfile -Sta -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TARGET%"
