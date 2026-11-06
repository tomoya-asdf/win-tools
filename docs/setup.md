# セットアップ

[README](../README.md)。以下はリポジトリのルートで実行します。登録は現在のユーザーが対象で、管理者権限は不要です。各ツールは登録せず直接実行できます。

## 起動ショートカットキー

スタートメニューに`Win Tools.lnk`を作成し、既定で **Ctrl + Alt + W** を割り当てます。

```powershell
./setup/Register-WinToolsShortcut.ps1

# キーを変更
./setup/Register-WinToolsShortcut.ps1 -HotKey 'CTRL+ALT+T'

# 登録内容の確認のみ / 解除
./setup/Register-WinToolsShortcut.ps1 -WhatIf
./setup/Register-WinToolsShortcut.ps1 -Unregister
```

通常権限で起動します。他のアプリとのキー競合は自動判定しません。リポジトリを移動した場合は再登録してください。

## 右クリックメニュー

```powershell
# 一時保存へコピー
./setup/Register-WinToolsMenus.ps1 -RegisterContextMenu

# Excelを読み取り専用で開く
./setup/Register-WinToolsMenus.ps1 -RegisterExcelReadOnly

# 起動キーと両メニューをまとめて登録（-WhatIfで事前確認）
./setup/Register-WinToolsMenus.ps1 -RegisterShortcut -RegisterContextMenu -RegisterExcelReadOnly
```

一時保存先は`%USERPROFILE%\一時保存フォルダ`です。変更する場合は`-TemporaryFolderPath 'D:\一時保存'`を付け、コピー・削除・監視にも同じ保存先を指定します。[一時保存の詳細](temporary-save.md)

Excelの対象は`.xlsx`・`.xlsm`・`.xls`・`.xlsb`・`.csv`で、`EXCEL.EXE /r`を使います。登録先は`HKCU\Software\Classes\SystemFileAssociations\<拡張子>\shell\OpenReadOnly`です。

## 右クリックメニューの解除

```powershell
./setup/Unregister-TemporarySaveMenu.ps1

# Excel読み取り専用メニューを解除
foreach ($extension in '.xlsx', '.xlsm', '.xls', '.xlsb', '.csv') {
  $key = "HKCU:\Software\Classes\SystemFileAssociations\$extension\shell\OpenReadOnly"
  if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse -Force }
}
```

登録・解除は`-WhatIf`に対応します。リポジトリや保存先を移動した場合は、メニューを再登録してください。

## 以前の構成からの移行

旧セットアップで一時保存先を`home/`へのJunctionにしていた場合だけ、次の対応が必要です。

1. 監視をトレイの`Exit`で終了します。
2. リンク先のデータをリポジトリ外へコピーし、内容を確認します。
3. プロファイル側のリンクだけを解除し、通常の保存先を作成してデータを戻します。リンク先を再帰削除しないでください。
4. メニューと監視のショートカットを現在のスクリプトへ登録し直します。

## 実行ポリシー

スクリプトの直接実行が制限される場合は、組織の設定を確認してください。必要に応じて`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`を使用します。
