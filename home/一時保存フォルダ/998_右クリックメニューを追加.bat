@echo off
setlocal

REM 管理者権限での実行チェック
net session >nul 2>&1
if %errorlevel% == 0 (
echo 管理者権限で実行されています。
) else (
echo 管理者権限で実行してください。
pause & exit /b 1
)

REM 実行するバッチファイルのパス（ここを必要に応じて変更）
set "BAT_PATH=%~dp0%\101_CopyToFolder.bat"

REM バッチファイルが存在するか確認
if not exist "%BAT_PATH%" (
    echo バッチファイルが見つかりません: %BAT_PATH%
    pause & exit /b 1
)

REM レジストリに右クリックメニューを追加（ファイル用）
reg add "HKCU\Software\Classes\*\shell\CopyToSelfFolder" /ve /d "コピーして開く（一時保存）" /f
reg add "HKCU\Software\Classes\*\shell\CopyToSelfFolder\command" /ve /d "\"%BAT_PATH%\" \"%%1\"" /f

REM レジストリに右クリックメニューを追加（フォルダ用）
reg add "HKCU\Software\Classes\Directory\shell\CopyToSelfFolder" /ve /d "コピーして開く（一時保存）" /f
reg add "HKCU\Software\Classes\Directory\shell\CopyToSelfFolder\command" /ve /d "\"%BAT_PATH%\" \"%%1\"" /f

echo.
echo 右クリックメニューを追加しました:
echo   表示名: コピーして開く（バッチ）
echo   実行: %BAT_PATH%

endlocal
pause