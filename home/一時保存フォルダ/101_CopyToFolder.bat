@echo off
setlocal

REM 引数チェック
if "%~1"=="" (
    echo コピーするファイルまたはフォルダのパスを引数で指定してください。
    pause & exit /b 1
)

REM 実行中のバッチファイルが存在するフォルダ
set "TARGET=%~dp0"
REM コピー元パス
set "SOURCE=%~1"
REM 日付取得（YYYYMMDD形式）
for /f "tokens=2 delims==" %%I in ('"wmic os get LocalDateTime /value"') do set "DATETIME=%%I"
set "DATE=%DATETIME:~0,8%"
REM コピー名（YYYYMMDD_元の名前）
set "COPYNAME=【一時保存】＜%DATE%＞%~nx1"

REM コピー元がフォルダかどうか判定
for %%F in ("%SOURCE%") do (
    if exist "%%~F\*" (
        REM フォルダ処理
        xcopy "%%~F" "%TARGET%\%COPYNAME%" /E /I /Y
        set "OPEN_PATH=%TARGET%\%COPYNAME%"
    ) else if exist "%%~F" (
        REM ファイル処理
        copy "%%~F" "%TARGET%\%COPYNAME%" /Y
        set "OPEN_PATH=%TARGET%\%COPYNAME%"
    ) else (
        echo コピー元が存在しません。
        pause & exit /b 1
    )
)
echo.
echo コピー完了:
echo   元: %SOURCE%
echo   先: %OPEN_PATH%
REM コピーしたファイルまたはフォルダを開く
start "" "%OPEN_PATH%"

endlocal