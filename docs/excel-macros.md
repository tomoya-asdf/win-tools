# Excelマクロツール

[概要に戻る](../README.md)。将来追加するマクロツールの配置先は `tools/excel-macros/` です。現時点では配置用の `.gitignore` だけを用意しています。

## 配置と起動

次のようにマクロのファイルを追加します。用途別のサブフォルダも利用できます。

```text
tools/excel-macros/
├─ .gitignore
├─ 集計ツール.xlsm
└─ 帳票/
   └─ 帳票作成.xlsm
```

対応する[Excelの形式](https://support.microsoft.com/en-us/excel/file-formats-that-are-supported-in-excel)は `.xlsm`・`.xlsb`・`.xlam`・`.xltm` です。拡張子の大文字小文字は区別しません。`.bas`・`.cls`・`.frm` はVBAソースとして配置できますが、単体で開いて実行するブックではないためメニューには表示しません。

`Start-WinTools.ps1` または `Start-WinTools.bat` を起動すると、「Excel: 帳票\帳票作成.xlsm」のように相対パスで一覧へ追加されます。メニューを再表示するたびに一覧を作り直すので、追加時の登録作業は不要です。

選択するとWindowsのファイル関連付けで開きます。対応するExcelと関連付けが必要です。個別のVBAプロシージャ名は指定せず、開いたブックのボタンや操作画面から利用する形です。ブックの起動時処理やマクロの有効・無効はExcel側の設定に従います。

## ツールと作業データの区別

Gitで管理するツール本体は `tools/excel-macros/` に置きます。このフォルダ内に限り、対応するマクロ形式を `.gitignore` の除外から戻しています。通常の業務ブックを置く他の場所では、Excelファイルの除外を継続します。

入力データや生成した帳票はリポジトリ外、またはマクロ用フォルダ内の `data/`・`input/`・`output/`・`local/` に保存します。これらのフォルダはGit管理とメニューの検索から除外します。`~$` で始まるExcelのロックファイル、`*.local.*`・`*.secrets.*`・`*.credentials.*` も対象外です。

隠し項目、名前が `.` で始まる項目、子のJunction・SymbolicLinkも検索しません。Gitに追加するツールブックには、個人情報・実際の業務データ・接続用の秘密情報が入っていないことを確認してください。入力用の `.xlsx` はこのフォルダ内でもGitの除外対象です。
