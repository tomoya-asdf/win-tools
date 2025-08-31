# win-tools (Windows dotfiles-like tools)

**概要**
- **目的:** `home` 配下のツールや設定をユーザープロファイルへリンク（疑似 dotfiles）。
- **右クリック:** 「一時保存」メニュー登録を支援。Excel の読み取り専用オープンも追加可能。
- **対応環境:** Windows 10/11、PowerShell 5.1/7+。

**初期セットアップ**
- **実行ポリシー:** `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
- **一時保存メニュー:** `./scripts/Install-WinTools.ps1 -RegisterContextMenu -Verbose`（管理者で実行）
- **Excel(読み取り専用):** `./scripts/Install-WinTools.ps1 -RegisterExcelReadOnly -Verbose`（管理者不要）

**主なコマンド**
- `./scripts/New-HomeLinks.ps1`: **リンク作成**（上書きなし）
- `./scripts/Install-WinTools.ps1 -Overwrite -BackupRoot "$env:USERPROFILE\win-tools_backup" -Verbose`: **バックアップして上書きリンク**

**右クリックメニュー**
- **一時保存:** `home/一時保存フォルダ/998_右クリックメニューを追加.bat` を昇格実行して登録。
- **Excel(読み取り専用):** `HKCU\Software\Classes\SystemFileAssociations\<ext>\shell\OpenReadOnly` に登録。
  - **対象拡張子:** `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`
  - **表示名:** `Excel(読み取り専用)で開く`
  - **コマンド:** `cmd.exe /c start "" excel /r "%1"`（Excel 実体を検出できればフルパス使用）
  - **アイコン:** `EXCEL.EXE,0`

**一時保存フォルダのツール**
- `101_CopyToFolder.bat`: 対象のファイル/フォルダを一時保存フォルダにコピー
  - 使い方: 右クリックメニュー（登録済みの場合）から実行、または `101_CopyToFolder.bat <パス>`
  - 動作: `【一時保存】YYYYMMDD_元名` の形式でコピーし、エクスプローラーで開く
  - 権限: 通常ユーザーで可
- `102_ClearCopiedItems.bat`: 一時保存フォルダ内のコピー品をまとめて削除
  - 使い方: ダブルクリックで実行（削除前に Y/N 確認）
  - 対象: `【一時保存】??????` で始まるファイル/フォルダ
  - 権限: 通常ユーザーで可
- `103_AutoRename.bat`: 自動リネーム監視をバックグラウンド起動（トレイ常駐）
  - 使い方: ダブルクリックで起動（非表示で起動、トレイから Exit で終了）
  - 動作: このフォルダ直下で新規作成された項目を `【一時保存】YYYYMMDD_` でリネーム
  - 自動起動: `shell:startup` にショートカットを配置
  - 権限: 通常ユーザーで可
- `400_RenameWatcher.ps1`: 監視本体（`103_AutoRename.bat` から起動）
  - 直接起動: `powershell -NoProfile -Sta -ExecutionPolicy Bypass -File "home\一時保存フォルダ\400_RenameWatcher.ps1"`
  - 注意: サブフォルダは対象外。`~` 始まり/`3桁_` 始まり/既に`【一時保存】`始まりはスキップ
- `998_右クリックメニューを追加.bat`: 一時保存メニューを登録（管理者）
  - 使い方: 管理者で実行。右クリックメニューに「コピーして自分フォルダへ（仮: 一時保存）」を追加
- `999_右クリックメニューを削除.bat`: 一時保存メニューを削除（管理者）
  - 使い方: 管理者で実行。登録したメニューを削除

**リンクの挙動**
- **ディレクトリ:** Junction 作成（管理者不要）
- **ファイル:** SymbolicLink → HardLink → 失敗時はコピーを順に試行

**文字コード（PowerShell）**
- **推奨形式:** UTF-8(BOM) + CRLF（`.ps1/.psm1/.psd1/.ps1xml`）
- **再チェックアウト:** もし以前に CP932 で取得していた場合は再取得
  - 例: `git rm --cached -r .` → `git reset --hard`

**注意事項**
- **HardLink:** 同一ドライブのみ。異なる場合や失敗時はコピー。
- **Symlink:** 開発者モードまたは管理者権限が必要な場合あり。

**アンインストール**
- 一時保存メニュー: 管理者で `./home/一時保存フォルダ/999_右クリックメニューを削除.bat` を実行。
- Excel(読み取り専用)メニュー: `HKCU` から該当キーを削除（管理者不要）。
  - コマンド: `$exts = '.xlsx','.xlsm','.xls','.xlsb','.csv'; foreach ($e in $exts) { $k = "HKCU:\\Software\\Classes\\SystemFileAssociations\\$e\\shell\\OpenReadOnly"; if (Test-Path $k) { Remove-Item -Recurse -Force $k } }`
- 作成したリンクの削除: `home` 直下名に対応する項目をユーザープロファイルから削除。
  - プレビュー: `Get-ChildItem ./home | % { Join-Path $env:USERPROFILE $_.Name } | % { $_ }`
  - 実行（リンク/ジャンクション優先で削除、コピーは手動判断推奨）:
    `Get-ChildItem ./home | % { $t = Join-Path $env:USERPROFILE $_.Name; if (Test-Path $t) { Remove-Item $t -Recurse -Force -WhatIf } }`
    問題なければ `-WhatIf` を外して実行してください。
- バックアップから復元: `-Overwrite -BackupRoot` を使っていた場合、`BackupRoot` 配下のタイムスタンプフォルダから元の場所へ戻してください。