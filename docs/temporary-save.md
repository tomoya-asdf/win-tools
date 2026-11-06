# 一時保存ツール

[概要に戻る](../README.md)。以下のコマンドはリポジトリのルートで実行します。

## コードと保存先

コピー・削除・監視の実装は `tools/temporary-save/`、右クリックメニューの登録・解除は `setup/` に置きます。各バッチは対応するPowerShellと同じベース名です。

コピー先・削除対象・監視先の既定は `%USERPROFILE%\一時保存フォルダ` です。各コマンドの `-FolderPath` で変更できます。コードの場所やカレントディレクトリを変えても、保存先は変わりません。コピー・監視の開始時に保存先がなければ作成します。

保存データはリポジトリ外に置きます。旧セットアップのJunctionがある場合は、[移行手順](setup.md#以前の構成からの移行)を確認してください。

## コピー

```powershell
./tools/temporary-save/Copy-TemporaryItem.ps1 -LiteralPath 'C:\Work\案件.xlsx' -Open

# 任意の保存先。実行前の確認には -WhatIf を追加
./tools/temporary-save/Copy-TemporaryItem.ps1 -LiteralPath 'C:\Work\案件フォルダ' -FolderPath 'D:\一時保存'
```

ファイルとフォルダを `【一時保存】＜YYYYMMDD＞元の名前` でコピーします。`-Open` は、ファイルなら保存したコピーを既定のアプリで開き、フォルダならコピーしたフォルダ自体をエクスプローラーで開きます。バッチ・右クリックメニュー・起動メニューからのコピーは `-Open` 付きで実行します。同名のコピーが存在する場合は上書きせずエラーにします。コピー元自身やその配下を保存先に指定することはできません。

`tools/temporary-save/Copy-TemporaryItem.bat` にファイルやフォルダをドラッグしても実行できます。日付の取得にはPowerShellを使い、WMICには依存しません。

## コピー品の削除

```powershell
# まず削除対象を確認
./tools/temporary-save/Clear-TemporaryFolder.ps1 -WhatIf

# 対象ごとに確認して削除
./tools/temporary-save/Clear-TemporaryFolder.ps1
```

保存先の直下にある、名前が `【一時保存】＜8桁の数字＞` で始まる項目だけが対象です。対象がフォルダならその内容も削除します。通常のフォルダの内部まで削除候補を探しません。JunctionやSymbolicLinkを含む候補は自動削除せずエラーにします。

ダブルクリック用の入口は `tools/temporary-save/Clear-TemporaryFolder.bat` です。別の保存先を使う場合は、コピー・監視と同じ `-FolderPath` を指定してください。

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

非表示で起動する入口は `tools/temporary-save/Start-TemporaryFolderWatcher.bat` です。自動起動したい場合は、このバッチへのショートカットを `shell:startup` に配置します。

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
