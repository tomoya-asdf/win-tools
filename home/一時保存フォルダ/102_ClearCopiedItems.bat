@echo off
setlocal

REM 実行中のバッチファイルが存在するフォルダ
set "TARGET=%~dp0"
set "TARGET_QUERY=%TARGET%【一時保存】＜????????＞*"

echo 以下のファイル・フォルダを削除します:
echo   対象: %TARGET_QUERY%

REM 削除確認（ユーザーに選択させる）
set /p CONFIRM="本当に削除しますか？ (Y/N): "
if /I not "%CONFIRM%"=="Y" (
    echo キャンセルされました。
    pause & exit /b 1
)

REM ファイル削除
for /f "delims=" %%F in ('dir "%TARGET_QUERY%" /b /s') do (
    if exist "%%F" (
        echo 削除中: %%~nxF
        del /q "%%F" 2>nul
        rmdir /s /q "%%F" 2>nul
    )
)
echo.
echo 削除完了しました。
endlocal
pause