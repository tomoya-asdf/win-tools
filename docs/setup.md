# セットアップ（メニュー登録）

[概要に戻る](../README.md)。以下のコマンドはリポジトリのルートから実行します。

`setup/` は、右クリックメニューの登録・解除を担当します。各ツールは直接実行でき、一時保存先はコピー・監視の開始時に作成します。プロファイルへのリンク作成や `home/` は不要です。

## 初期設定

必要に応じて、現在のユーザーのPowerShell実行ポリシーを設定します。

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## 一時保存の右クリックメニュー

現在のユーザーのメニューに登録します。管理者権限は不要です。

```powershell
./setup/Register-WinToolsMenus.ps1 -RegisterContextMenu -Verbose

# 一時保存メニュー単独でも登録できる
./setup/Register-TemporarySaveMenu.ps1

# 別の保存先をメニューに固定する場合
./setup/Register-TemporarySaveMenu.ps1 -FolderPath 'D:\一時保存'
```

`Register-WinToolsMenus.ps1` は `Register-TemporarySaveMenu.ps1` に登録を委譲します。前者の保存先オプションは `-TemporaryFolderPath` で、既定は `%USERPROFILE%\一時保存フォルダ` です。コピー・削除・監視も同じ保存先を指定して利用してください。詳しくは[一時保存ツール](temporary-save.md)を参照してください。

## Excelを読み取り専用で開くメニュー

```powershell
./setup/Register-WinToolsMenus.ps1 -RegisterExcelReadOnly -Verbose

# 両方のメニューを登録
./setup/Register-WinToolsMenus.ps1 -RegisterContextMenu -RegisterExcelReadOnly

# 登録内容を確認してから実行する場合
./setup/Register-WinToolsMenus.ps1 -RegisterContextMenu -RegisterExcelReadOnly -WhatIf
```

対象は `.xlsx`、`.xlsm`、`.xls`、`.xlsb`、`.csv` です。`HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\OpenReadOnly` に登録します。管理者権限は不要です。表示名は「Excel(読み取り専用)で開く」、実行コマンドはExcelの `/r` オプションを使用します。

Excelの実体を検出できればフルパスを使い、検出できない場合は `cmd.exe /c start "" excel /r "%1"` を登録します。メニューのアイコンは `EXCEL.EXE,0` を使用します。登録オプションなしで起動した場合は、使い方を表示して終了します。

## 解除

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

## 以前の構成からの移行

`home/` のテンプレートと `New-HomeLinks.ps1` は廃止しました。リンクを使っていなければ追加の移行は不要です。以前のセットアップで `%USERPROFILE%\一時保存フォルダ` を旧 `home/一時保存フォルダ` へのJunctionにしていた場合は、次の順でデータを独立させてください。

1. 自動リネーム監視をトレイの `Exit` から終了します。
2. 旧リンク先の保存データを、リポジトリ外の別フォルダへコピーし、内容を確認します。フォルダ比較の `-CompareBy Hash` も利用できます。
3. プロファイル側のリンクを確認して解除します。リンク先の保存データをまとめて削除する操作は避けてください。
4. プロファイル直下に通常の `一時保存フォルダ` を作成し、退避したデータを戻します。別の保存先を使う場合は各ツールの `-FolderPath` を指定します。
5. `Register-TemporarySaveMenu.ps1` でメニューを再登録し、保存先と監視先を確認します。

既存の保存データがある環境では、リポジトリ内に残った旧 `home/` はデータ退避と確認後に整理してください。旧データが再作成されてもGitに混ざらないよう、ルートの `.gitignore` で `/home/` を除外しています。

旧 `Install-WinTools.ps1` は `Register-WinToolsMenus.ps1` に変更しました。`-RegisterContextMenu`・`-RegisterExcelReadOnly`・`-TemporaryFolderPath` は継続します。リンク作成用の `-MakeLinks`・`-SourceRoot`・`-TargetRoot`・`-Overwrite`・`-BackupRoot` は廃止したため、古い呼び出しから取り除いてください。

## 文字コード

新規・変更する `.ps1` / `.psm1` / `.psd1` / `.ps1xml` はUTF-8 BOM付き・CRLFで保存します。起動用バッチはASCII・CRLFで保存し、日本語の表示や処理はPowerShell側で行います。バッチと対応するPowerShellは同じベース名にします。
