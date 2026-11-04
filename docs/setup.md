# セットアップ

[概要に戻る](../README.md)。以下のコマンドはリポジトリのルートから実行します。

`setup/` は、ユーザー環境への配置とメニュー登録を担当します。ログ表示ツールだけを使う場合、セットアップは不要です。

## 初期設定

必要に応じて、現在のユーザーのPowerShell実行ポリシーを設定します。

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## ユーザープロファイルへのリンク

`home/` 直下の各項目を、同名で `%USERPROFILE%` 直下へリンクします。現在の対象はデータ保存用の `一時保存フォルダ` です。実行コードは含みません。`home/` 自体をプロファイルへリンクする処理ではありません。一時保存ツールを直接実行すればプロファイル直下に保存先を作成するので、リンクの利用は任意です。リンクを作成した場合、データの実体はリポジトリの `home/一時保存フォルダ` に置かれます。

```powershell
# まず実行内容を確認（変更しない）
./setup/New-HomeLinks.ps1 -WhatIf

# 既存の同名項目を上書きせずリンクを作成
./setup/New-HomeLinks.ps1 -Verbose

# 既存項目を退避してからリンクを作り直す
./setup/Install-WinTools.ps1 -Overwrite -BackupRoot "$env:USERPROFILE\win-tools_backup" -Verbose
```

- ディレクトリはJunctionを作成します（通常、管理者不要）。
- ファイルはSymbolicLink、HardLink、コピーの順に試します。SymbolicLinkには開発者モードや管理者権限が必要な場合があります。HardLinkは同じドライブ内でのみ作成できます。
- `-Overwrite` のみを指定すると既存項目を削除するため、内容を残す場合は必ず `-BackupRoot` も指定してください。
- `-SourceRoot` はリンクする内容のルート（既定はリポジトリの `home/`）、`-TargetRoot` はリンク先のルート（既定は `%USERPROFILE%`）を指定します。
- JunctionやSymbolicLinkはリポジトリ内を参照します。利用中はリポジトリや `home/` を移動しないでください。HardLinkやコピーへフォールバックしたファイルは、後からのファイル置換・更新がリンク先に反映されない場合があります。

## 一時保存の右クリックメニュー

現在のユーザーのメニューに登録するので、管理者権限は不要です。既定ではリンク作成も行います。

```powershell
./setup/Install-WinTools.ps1 -RegisterContextMenu -Verbose

# リンク作成なしでメニューだけ登録する場合
./setup/Register-TemporarySaveMenu.ps1

# 別の保存先をメニューに固定する場合
./setup/Register-TemporarySaveMenu.ps1 -FolderPath 'D:\一時保存'
```

メニューの登録先や操作方法は[一時保存ツール](temporary-save.md)を参照してください。`Install-WinTools.ps1` は `Register-TemporarySaveMenu.ps1` に登録を委譲します。前者では `-TemporaryFolderPath` で保存先を指定できます（既定は `TargetRoot/一時保存フォルダ`）。既存メニューも再実行で新しいツールのパスへ更新できます。

## Excelを読み取り専用で開くメニュー

管理者権限は不要です。

```powershell
./setup/Install-WinTools.ps1 -RegisterExcelReadOnly -Verbose

# リンク作成を行わず、メニュー登録だけを実行
./setup/Install-WinTools.ps1 -MakeLinks:$false -RegisterExcelReadOnly -Verbose
```

対象は `.xlsx`、`.xlsm`、`.xls`、`.xlsb`、`.csv` です。`HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\OpenReadOnly` に登録します。表示名は「Excel(読み取り専用)で開く」、実行コマンドはExcelの `/r` オプションを使用します。

Excelの実体を検出できればフルパスを使い、検出できない場合は `cmd.exe /c start "" excel /r "%1"` を登録します。メニューのアイコンは `EXCEL.EXE,0` を使用します。

## 解除と復元

- 一時保存メニューは `./setup/Unregister-TemporarySaveMenu.ps1` で解除します（管理者不要、`-WhatIf` 対応）。旧バッチと同じキーを使用するため、旧メニューも解除できます。
- Excelのメニューは、以下のコマンドで今回登録したキーを削除します。

```powershell
$extensions = '.xlsx', '.xlsm', '.xls', '.xlsb', '.csv'
foreach ($extension in $extensions) {
  $key = "HKCU:\Software\Classes\SystemFileAssociations\$extension\shell\OpenReadOnly"
  if (Test-Path -LiteralPath $key) {
    Remove-Item -LiteralPath $key -Recurse -Force
  }
}
```

リンクを解除する場合は、自動リネーム監視をトレイの `Exit` から終了し、リンク先に必要なデータがないか確認してください。Junction配下のデータはリポジトリ側にも存在するため、内容をまとめて削除する操作に注意が必要です。リンクとコピーされた実体を区別したうえで、プロファイル直下の対象項目だけを解除してください。

`-BackupRoot` で退避した項目は、バックアップ先のタイムスタンプ付きフォルダにあります。リンクの解除後、元の場所に戻してください。

## 文字コード

新規・変更する `.ps1` / `.psm1` / `.psd1` / `.ps1xml` はUTF-8 BOM付き・CRLFで保存します。一時保存の起動用バッチはASCII・CRLFで保存し、日本語の表示や処理はPowerShell側で行います。
