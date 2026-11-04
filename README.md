# win-tools

Windowsで使う作業ツールと、ユーザー環境を整えるセットアップをまとめたリポジトリです。Windows 10/11、PowerShell 5.1/7+を対象にしています。

## フォルダ構成

```text
win-tools/
├─ setup/                         # インストール・環境設定
│  ├─ Install-WinTools.ps1         # リンク作成とメニュー登録の入口
│  ├─ New-HomeLinks.ps1            # home/ の内容をユーザープロファイルへリンク
│  ├─ Register-TemporarySaveMenu.ps1
│  └─ Unregister-TemporarySaveMenu.ps1
├─ tools/                         # リポジトリから直接実行するツール
│  ├─ logs/                       # ログ表示と専用の内部実装
│  └─ temporary-save/             # 一時保存のコピー・削除・監視
│     ├─ Copy-ToTemporaryFolder.ps1
│     ├─ Clear-TemporaryFolder.ps1
│     ├─ Start-TemporaryFolderWatcher.ps1
│     ├─ 101_～103_*.bat           # ダブルクリック・ドラッグ操作の入口
│     └─ lib/TemporarySave.psm1
├─ home/                          # ユーザープロファイルに配置する内容
│  └─ 一時保存フォルダ/.gitignore    # データ保存先のテンプレート（コードなし）
├─ docs/                          # 機能別の使い方・制約・解除方法
│  ├─ setup.md
│  ├─ logs.md
│  └─ temporary-save.md
└─ tests/                         # 検証用コード
   ├─ Setup.Tests.ps1
   ├─ logs/ActivityLogs.Tests.ps1
   └─ temporary-save/TemporarySave.Tests.ps1
```

## できること

以下のコマンドはリポジトリのルートから実行します。

| 用途 | 実行するもの | 詳細 |
| --- | --- | --- |
| ユーザープロファイルへのリンク作成 | `./setup/New-HomeLinks.ps1` | [セットアップ](docs/setup.md) |
| 一時保存・Excel読み取り専用のメニュー登録 | `./setup/Install-WinTools.ps1` に登録オプションを指定 | [セットアップ](docs/setup.md) |
| ファイルやフォルダの一時保存・自動改名 | `tools/temporary-save/` 内のPowerShell・バッチ | [一時保存ツール](docs/temporary-save.md) |
| フォルダの作成・削除・名前変更などの履歴 | `./tools/logs/Get-RecentFolderActivity.ps1`（管理者） | [ログ表示](docs/logs.md) |
| エラー・再起動・ログオン・更新などの履歴 | `./tools/logs/Get-WorkEventLog.ps1` | [ログ表示](docs/logs.md) |

ログ表示・一時保存ツールはセットアップせずに利用できます。一時保存先の既定は `%USERPROFILE%\一時保存フォルダ` です。

```powershell
./tools/logs/Get-WorkEventLog.ps1 -Category Restart,WindowsUpdate

# 管理者として実行。既存のNTFS USNジャーナルを読み取る
./tools/logs/Get-RecentFolderActivity.ps1 |
  Format-Table TimeCreated, Operation, Name, PathHint -AutoSize -Wrap
```

## ファイルを追加するときのルール

- 普段使うツールは `tools/<用途>/` に配置します。専用の補助コードは同じ機能の `lib/` にまとめます。
- インストール・登録・解除など、環境を設定する処理は `setup/` に配置します。
- `home/` は `%USERPROFILE%` に配置するデータ用ディレクトリや設定を表す場所です。直下の名前がリンク先の名前になります。実行するコードは `tools/`、登録・解除の処理は `setup/` に置きます。
- 詳しい使い方は `docs/` にまとめ、READMEには機能の一覧と入口を載せます。
- テストは `tests/` に置き、用途ごとのテストが増えたらツールと対応するサブフォルダにまとめます。
- 新規・変更するPowerShellファイルはUTF-8 BOM付き・CRLFにします。保存データや実行結果はGitに追加せず、一時保存フォルダや任意のリポジトリ外の場所に保存してください。

保存先・監視先はコードの置き場所から独立しており、`-FolderPath` で変更できます。`home/一時保存フォルダ` に既存の保存データがある場合は、そのまま保持します。同フォルダの `.gitignore` はユーザーデータをGitから除外します。

## 以前の構成からの変更

`home/` と既存の保存データへのリンクは維持します。一時保存の旧監視を終了し、登録済みの右クリックメニューは `./setup/Register-TemporarySaveMenu.ps1` で再登録してください。起動ショートカットと手順書のパスも更新します。詳しくは[一時保存ツールの移行手順](docs/temporary-save.md#以前の構成からの移行)を参照してください。

| 以前 | 現在 |
| --- | --- |
| `scripts/Install-WinTools.ps1` | `setup/Install-WinTools.ps1` |
| `scripts/New-HomeLinks.ps1` | `setup/New-HomeLinks.ps1` |
| `scripts/Get-RecentFolderActivity.ps1` | `tools/logs/Get-RecentFolderActivity.ps1` |
| `scripts/Get-WorkEventLog.ps1` | `tools/logs/Get-WorkEventLog.ps1` |
| `home/一時保存フォルダ/101_CopyToFolder.bat` | `tools/temporary-save/101_CopyToFolder.bat` |
| `home/一時保存フォルダ/102_ClearCopiedItems.bat` | `tools/temporary-save/102_ClearCopiedItems.bat` |
| `home/一時保存フォルダ/103_AutoRename.bat` | `tools/temporary-save/103_AutoRename.bat` |
| `home/一時保存フォルダ/400_RenameWatcher.ps1` | `tools/temporary-save/Start-TemporaryFolderWatcher.ps1` |
| `home/一時保存フォルダ/998_右クリックメニューを追加.bat` | `setup/Register-TemporarySaveMenu.ps1` |
| `home/一時保存フォルダ/999_右クリックメニューを削除.bat` | `setup/Unregister-TemporarySaveMenu.ps1` |

## 検証

```powershell
# Pesterが必要（3.4で動作確認）。セットアップのテストも環境を変更せず実行
Invoke-Pester ./tests
```

USNの実ドライブ読み取りは管理者として別途確認してください。
