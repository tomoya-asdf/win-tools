# win-tools

Windows用のツールと起動メニューです。Windows 10/11、PowerShell 5.1/7+が対象です。Excel機能にはデスクトップ版Excelが必要です。

## 起動

`Start-WinTools.bat`をダブルクリックするか、リポジトリのルートで実行します。

```powershell
./Start-WinTools.ps1

# Ctrl + Alt + Wを登録（初回のみ）
./setup/Register-WinToolsShortcut.ps1
```

番号で選択し、`0`で終了します。機能の入力前に画面をクリアし、結果はEnterを押すまで表示します。

## アプリを追加

初回起動時に作成される`apps/`に、ショートカット・exe・URLリンク・文書を置きます。フォルダは子メニューになります。

```text
apps/                   # 全体をGit管理から除外
├─ ブラウザ.lnk
├─ ツール.exe
└─ 仕事/
   ├─ 社内サイト.url
   └─ 手順書.pdf
```

追加・削除はメニューの再表示時に反映されます。ファイルは既定のアプリで開き、子メニューの`0`で一つ上へ戻ります。アプリ本体を移動したくない場合は、ショートカットを置いてください。

## 機能と使い方

| 機能 | ドキュメント |
| --- | --- |
| 起動メニュー・追加アプリ・実行履歴 | [起動メニュー](docs/launcher.md) |
| 起動キー・右クリックメニューの登録と解除 | [セットアップ](docs/setup.md) |
| フォルダの不足項目・サイズ・内容の比較 | [フォルダ比較](docs/files.md) |
| ファイル一覧・Excelシート一覧 | [Excelマクロ](docs/excel-macros.md) |
| 一時保存・コピー品の削除・自動改名 | [一時保存](docs/temporary-save.md) |
| Windowsイベント・フォルダ変更の履歴 | [ログ表示](docs/logs.md) |

実行履歴は`%LOCALAPPDATA%\WinTools\History`、一時保存先は`%USERPROFILE%\一時保存フォルダ`です。

## 開発・検証

実装は`tools/`、環境への登録は`setup/`、検証コードは`tests/`に配置します。

```powershell
# Pester 3.4で動作確認
Invoke-Pester ./tests
```

[開発ルール](docs/development.md)に命名・文字コード・Git管理の方針をまとめています。
