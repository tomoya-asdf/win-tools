# win-tools

Windowsで使う作業ツールと、ユーザー環境を整えるセットアップをまとめたリポジトリです。Windows 10/11、PowerShell 5.1/7+を対象にしています。

## フォルダ構成

```text
win-tools/
├─ Start-WinTools.ps1             # 番号選択の起動メニュー
├─ Start-WinTools.bat             # ダブルクリック用の入口
├─ setup/                         # メニュー登録・解除
│  ├─ Register-WinToolsMenus.ps1   # 一時保存・Excelのメニュー登録
│  ├─ Register-TemporarySaveMenu.ps1
│  └─ Unregister-TemporarySaveMenu.ps1
├─ tools/                         # リポジトリから直接実行するツール
│  ├─ logs/                       # ログ表示と専用の内部実装
│  ├─ files/Compare-Folders.ps1    # フォルダ比較
│  ├─ launcher/lib/               # メニューの一覧作成・起動処理
│  ├─ excel-macros/               # 追加したマクロブックをメニューで自動検出
│  └─ temporary-save/             # 一時保存のコピー・削除・監視
│     ├─ Copy-TemporaryItem.ps1 / .bat
│     ├─ Clear-TemporaryFolder.ps1 / .bat
│     ├─ Start-TemporaryFolderWatcher.ps1 / .bat
│     └─ lib/TemporarySave.psm1
├─ docs/                          # 機能別の使い方・制約・解除方法
│  ├─ setup.md
│  ├─ logs.md
│  ├─ launcher.md
│  ├─ files.md
│  ├─ excel-macros.md
│  └─ temporary-save.md
└─ tests/                         # 検証用コード
   ├─ Setup.Tests.ps1
   ├─ logs/ActivityLogs.Tests.ps1
   ├─ files/Compare-Folders.Tests.ps1
   ├─ launcher/WinToolsMenu.Tests.ps1
   └─ temporary-save/TemporarySave.Tests.ps1
```

## できること

以下のコマンドはリポジトリのルートから実行します。

| 用途 | 実行するもの | 詳細 |
| --- | --- | --- |
| 番号からツール・Excelマクロを選んで起動 | `./Start-WinTools.ps1` または `Start-WinTools.bat` | [起動メニュー](docs/launcher.md) |
| フォルダ内の不足項目・サイズ差・内容差を確認 | `./tools/files/Compare-Folders.ps1` | [フォルダ比較](docs/files.md) |
| Excelマクロツールの追加 | `tools/excel-macros/` に対応ファイルを配置 | [Excelマクロツール](docs/excel-macros.md) |
| 一時保存・Excel読み取り専用のメニュー登録 | `./setup/Register-WinToolsMenus.ps1` に登録オプションを指定 | [セットアップ](docs/setup.md) |
| ファイルやフォルダの一時保存・自動改名 | `tools/temporary-save/` 内のPowerShell・バッチ | [一時保存ツール](docs/temporary-save.md) |
| フォルダの作成・削除・名前変更などの履歴 | `./tools/logs/Get-RecentFolderActivity.ps1`（管理者） | [ログ表示](docs/logs.md) |
| エラー・再起動・ログオン・更新などの履歴 | `./tools/logs/Get-WorkEventLog.ps1` | [ログ表示](docs/logs.md) |

ログ表示・一時保存ツールはセットアップせずに利用できます。一時保存先の既定は `%USERPROFILE%\一時保存フォルダ` です。

```powershell
./Start-WinTools.ps1

./tools/logs/Get-WorkEventLog.ps1 -Category Restart,WindowsUpdate

# 管理者として実行。既存のNTFS USNジャーナルを読み取る
./tools/logs/Get-RecentFolderActivity.ps1 |
  Format-Table TimeCreated, Operation, Name, PathHint -AutoSize -Wrap
```

## ファイルを追加するときのルール

- 普段使うツールは `tools/<用途>/` に配置します。専用の補助コードは同じ機能の `lib/` にまとめます。
- Excelマクロツールのブックは `tools/excel-macros/` に置きます。対応形式は自動で起動メニューに追加されます。入力・出力データはツール本体と分けます。
- インストール・登録・解除など、環境を設定する処理は `setup/` に配置します。
- 詳しい使い方は `docs/` にまとめ、READMEには機能の一覧と入口を載せます。
- テストは `tests/` に置き、用途ごとのテストが増えたらツールと対応するサブフォルダにまとめます。
- 新規・変更するPowerShellファイルはUTF-8 BOM付き・CRLFにします。保存データや実行結果はGitに追加せず、一時保存フォルダや任意のリポジトリ外の場所に保存してください。

`.gitignore` はログ出力（CSV・EVTXなど）、診断ダンプ、業務文書、ローカル設定、秘密鍵・認証設定を除外します。出力用の `exports/`・`reports/`・`output/` なども対象です。既に追跡済みのファイルやソース内に書いた情報は除外されないため、コミット前に差分を確認してください。

例外として `tools/excel-macros/` 内のマクロツール本体（`.xlsm`・`.xlsb`・`.xlam`・`.xltm`）はGitで管理できます。同フォルダ内の `data/`・`input/`・`output/`・`local/`、ロックファイルやローカル設定は引き続き除外します。

保存先・監視先は `%USERPROFILE%\一時保存フォルダ` が既定で、`-FolderPath` で変更できます。コピー・監視の開始時にフォルダが作成されます。保存データはリポジトリの外に置きます。

## 命名規則

- フォルダ・文書は英小文字とハイフンを使います（`temporary-save/`、`excel-macros.md` など）。
- 実行するPowerShellは `動詞-対象.ps1` とし、単語の先頭を大文字にします（`Compare-Folders.ps1` など）。番号の接頭辞は付けません。
- バッチは呼び出すPowerShellと同じベース名にします（`Copy-TemporaryItem.bat` → `Copy-TemporaryItem.ps1`）。
- 補助モジュール・クラスは `TemporarySave.psm1` のように内容を表す名前にし、テストは `機能名.Tests.ps1` とします。

## 以前の構成からの変更

`home/` とリンク作成機能は廃止しました。旧セットアップで保存先をリンクしていた場合は、[セットアップの移行手順](docs/setup.md#以前の構成からの移行)を確認してください。番号付きバッチはPowerShellと同じ名前へ変更したため、[一時保存ツールの移行手順](docs/temporary-save.md#以前の構成からの移行)に沿ってショートカットとメニューを更新してください。

## 検証

```powershell
# Pesterが必要（3.4で動作確認）。セットアップのテストも環境を変更せず実行
Invoke-Pester ./tests
```

USNの実ドライブ読み取りは管理者として別途確認してください。
