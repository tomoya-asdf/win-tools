# フォルダ比較

[概要に戻る](../README.md)。コピー元・コピー先を相対パスで比較します。ファイルのコピー・削除・同期は行いません。

```powershell
# 相対パス・種類・サイズで比較
./tools/files/Compare-Folders.ps1 -LeftPath 'C:\Work\original' -RightPath 'D:\Backup' |
  Format-Table Status, RelativePath, LeftLength, RightLength, Note -AutoSize -Wrap

# SHA-256で内容も比較し、差分と比較不能な項目を表示
./tools/files/Compare-Folders.ps1 -LeftPath 'C:\Work\original' -RightPath 'D:\Backup' -CompareBy Hash -DifferencesOnly

# 結果をCSV保存（gitignore対象の出力先）
./tools/files/Compare-Folders.ps1 -LeftPath 'C:\Work\original' -RightPath 'D:\Backup' -CompareBy Hash |
  Export-Csv "$env:TEMP\folder-comparison.csv" -NoTypeInformation -Encoding UTF8
```

`Start-WinTools.ps1` の「フォルダ比較」からも実行できます。

## 比較方法

- 既定の `-CompareBy Size` は同じ相対パスの種類・ファイルサイズを比較します。同じサイズでも内容は異なる可能性があります。
- `-CompareBy Hash` は同じサイズのファイルについて、[Get-FileHash](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-filehash?view=powershell-5.1) のSHA-256で内容を比較します。サイズが異なるファイルはハッシュを計算せず差分にします。ハッシュ比較はファイルを読み取るため、大きいファイルでは時間がかかります。
- 名前はWindowsの通常のパスと同じく、大文字小文字を区別せず照合します。更新日時は結果に含めますが、一致・不一致の判定条件にはしません。
- 隠し項目、空フォルダ、サブフォルダ内も比較対象です。パス中の `[`・`]` などはワイルドカードとして扱いません。

## 結果のStatus

| Status | 意味 |
| --- | --- |
| `OnlyLeft` / `OnlyRight` | 左側 / 右側だけに存在する |
| `TypeMismatch` | 同じ相対パスが片側はファイル、反対側はフォルダ |
| `DifferentSize` | ファイルサイズが異なる |
| `DifferentContent` | 同じサイズだがSHA-256が異なる |
| `SameSize` | サイズ一致。内容は未確認 |
| `SameContent` | SHA-256一致 |
| `Directory` | 両側に同じ相対パスのフォルダが存在する |
| `SkippedLink` | 子のJunction・SymbolicLinkなどを検出したため比較対象から除外 |
| `ReadError` | フォルダの列挙やファイルの読み取りに失敗した、またはハッシュ比較中に変更を検出した |
| `NotCompared` | 反対側の親が列挙不可・リンクのため、存在や内容を比較できない |

`-DifferencesOnly` は `SameSize`・`SameContent`・`Directory` を省略します。リンクや読み取りエラーは残るので、差分件数だけで完全な一致とは判断せず、Statusも確認してください。空フォルダ同士は結果0件です。

リンクは子の探索時にたどりません。明示的に指定した比較ルートがJunctionの場合は、その保存先の直下から比較します。アクセス権、所有者、代替データストリーム、リンク先、ACLの比較は対象外です。Windowsで大文字小文字を区別するよう設定したディレクトリには対応していません。

実行中の変更を固定するスナップショットではありません。ハッシュの前後でサイズ・更新日時の変更を検出した場合は `ReadError` にしますが、書き込みが続くフォルダでは、作業が落ち着いた後に再度比較してください。
