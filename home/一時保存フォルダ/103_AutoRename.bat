@echo off
setlocal

REM 名称監視スクリプト
set "TARGET=%~dp0\400_RenameWatcher.ps1"

powershell -ExecutionPolicy Bypass -File %TARGET%