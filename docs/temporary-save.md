# 一時保存ツール

[概要に戻る](../README.md)。以下のコマンドはリポジトリのルートで実行します。

## コードと保存先

コピー・削除・監視の実装は `tools/temporary-save/`、右クリックメニューの登録・解除は `setup/` に置きます。`home/一時保存フォルダ` には実行コードを置きません。

コピー先・削除対象・監視先の既定は `%USERPROFILE%\一時保存フォルダ` です。各コマンドの `-FolderPath` で変更できます。コードの場所やカレントディレクトリを変えても、保存先は変わりません。コピー・監視の開始時に保存先がなければ作成します。

リンクのセットアップは必須ではありません。以前のセットアップでプロファイルの保存先が `home/一時保存フォルダ` へのJunctionになっている場合は、そのまま利用できます。保存データや既存のリンクは自動移行・削除しません。

## コピー

```powershell
./tools/temporary-save/Copy-ToTemporaryFolder.ps1 -LiteralPath 'C:\Work\案件.xlsx' -Open

# 任意の保存先。実行前の確認には -WhatIf を追加
./tools/temporary-save/Copy-ToTemporaryFolder.ps1 -LiteralPath 'C:\Work\案件フォルダ' -FolderPath 'D:\一時保存'
```

ファイルとフォルダを `【一時保存】＜YYYYMMDD＞元の名前` でコピーします。`-Open` はコピー先の項目をエクスプローラーで選択表示します。同名のコピーが存在する場合は上書きせずエラーにします。コピー元自身やその配下を保存先に指定することはできません。

`tools/temporary-save/101_CopyToFolder.bat` にファイルやフォルダをドラッグしても実行できます。日付の取得にはPowerShellを使い、WMICには依存しません。

## コピー品の削除

```powershell
# まず削除対象を確認
./tools/temporary-save/Clear-TemporaryFolder.ps1 -WhatIf

# 対象ごとに確認して削除
./tools/temporary-save/Clear-TemporaryFolder.ps1
```

保存先の直下にある、名前が `【一時保存】＜8桁の数字＞` で始まる項目だけが対象です。対象がフォルダならその内容も削除します。通常のフォルダの内部まで削除候補を探しません。JunctionやSymbolicLinkを含む候補は自動削除せずエラーにします。

ダブルクリック用の入口は `tools/temporary-save/102_ClearCopiedItems.bat` です。別の保存先を使う場合は、コピー・監視と同じ `-FolderPath` を指定してください。

## 自動リネーム監視

```powershell
# トレイ常駐。トレイアイコンの Exit から終了
./tools/temporary-save/Start-TemporaryFolderWatcher.ps1

# 別の保存先を監視
./tools/temporary-save/Start-TemporaryFolderWatcher.ps1 -FolderPath 'D:\一時保存'

# コンソールで監視。Ctrl+Cで終了（DurationSeconds指定で自動終了も可能）
./tools/temporary-save/Start-TemporaryFolderWatcher.ps1 -NoTray
```

保存先の直下に新規作成されたファイル・フォルダに、コピーと同じ日付付きの名前を付けます。既存項目、サブフォルダの内部、`~` 始まり、`3桁_` 始まり、既に `【一時保存】` で始まる名前は対象外です。名前が競合した場合は上書きしません。

作成通知の約1秒後に改名を試み、使用中などで失敗した場合は最大5回試行します。継続して失敗した項目や、イベントバッファがあふれた場合は警告を出します。非表示での実行時に警告を確認したい場合は、`-NoTray` で起動してください。同じ保存先を複数の監視プロセスで同時に監視しないでください。

非表示で起動する入口は `tools/temporary-save/103_AutoRename.bat` です。自動起動したい場合は、このバッチへのショートカットを `shell:startup` に配置します。

## 右クリックメニュー

```powershell
# 現在のユーザーへ登録・再登録（管理者不要）
./setup/Register-TemporarySaveMenu.ps1

# 任意の保存先を使用
./setup/Register-TemporarySaveMenu.ps1 -FolderPath 'D:\一時保存'

# 解除
./setup/Unregister-TemporarySaveMenu.ps1
```

登録・解除とも `-WhatIf` に対応します。登録先は `HKCU\Software\Classes\*\shell\CopyToSelfFolder` と `HKCU\Software\Classes\Directory\shell\CopyToSelfFolder` です。「コピーして開く（一時保存）」から、新しいコピー用PowerShellを実行します。

メニューにはコードと保存先のパスを登録するため、リポジトリや保存先を移動した場合は再登録してください。登録だけでは監視やスタートアップ設定は行いません。まとめてセットアップする方法は[セットアップ](setup.md)を参照してください。

## 以前の構成からの移行

1. 旧 `400_RenameWatcher.ps1` が動いていれば、トレイアイコンの `Exit` で終了します。
2. 保存データと既存のJunctionはそのまま使います。今回の変更では自動移動・削除しません。
3. `./setup/Register-TemporarySaveMenu.ps1` を実行し、旧メニューの参照先を更新します。別の保存先を使っていた場合は `-FolderPath` でその場所を指定します。
4. バッチへのショートカットを `tools/temporary-save/` 内の同名バッチに更新します。スタートアップに登録していた `103_AutoRename.bat` のショートカットも更新します。
5. 新しい `103_AutoRename.bat` または `Start-TemporaryFolderWatcher.ps1` で監視を開始します。保存先を変更する場合は、コピー・削除・監視・メニューで同じ場所を指定します。

旧 `998_右クリックメニューを追加.bat` と `999_右クリックメニューを削除.bat` は、それぞれ `setup/Register-TemporarySaveMenu.ps1` と `setup/Unregister-TemporarySaveMenu.ps1` に置き換えました。
