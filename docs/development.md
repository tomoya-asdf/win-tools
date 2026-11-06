# 開発ルール

[README](../README.md)

## 配置・命名

- 実行ツールは`tools/<用途>/`、補助コードはその`lib/`へ配置します。
- 環境への登録・解除は`setup/`、説明は`docs/`、検証は`tests/`にまとめます。
- フォルダ・文書名は英小文字とハイフン、PowerShellは`動詞-対象.ps1`とします。
- バッチは呼び出すPowerShellと同じベース名にします。
- PowerShellはUTF-8 BOM付き・CRLF、バッチはASCII・CRLFで保存します。

## Git管理

`apps/`は個人用の登録先で、全体を除外します。ログ・生成物・業務データ・認証情報もコミットしません。保存先にはリポジトリ外、または除外済みの`output/`・`local/`などを使用します。

共有するExcelマクロ本体は`tools/excel-macros/`で管理できます。その配下の`data/`・`input/`・`output/`・`local/`は除外します。マクロには実際の業務データや秘密情報を入れないでください。

既に追跡されているファイルは`.gitignore`で除外できないため、コミット前に差分を確認します。

## 検証

```powershell
# Pester 3.4で動作確認。登録系テストも実環境を変更しない
Invoke-Pester ./tests
```

実Excelの検証は[Excelマクロ](excel-macros.md#ソースと再生成)を参照してください。USNの実ドライブ読み取りは管理者として別途確認します。
