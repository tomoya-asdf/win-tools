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

echo 右クリックメニューを削除しています...
REM ファイル用の右クリックメニュー削除
reg delete "HKCU\Software\Classes\*\shell\CopyToSelfFolder" /f
REM フォルダ用の右クリックメニュー削除
reg delete "HKCU\Software\Classes\Directory\shell\CopyToSelfFolder" /f

echo.
echo 削除が完了しました。
endlocal
pause
