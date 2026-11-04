# ログ表示ツール

[概要に戻る](../README.md)。以下のコマンドはリポジトリのルートで実行します。

## 直近のフォルダ変更（既存のUSNジャーナル）

管理者として起動したPowerShellで、リポジトリのルートから実行します。追加インストールは不要です。`tools/logs` 配下のため、`New-HomeLinks.ps1` のリンク対象には含まれません。

```powershell
# システムドライブの過去24時間、最新100件
./tools/logs/Get-RecentFolderActivity.ps1 |
  Format-Table TimeCreated, Operation, Name, PathHint -AutoSize -Wrap

# Dドライブの過去7日間。「案件」を名前に含むフォルダを最新200件
./tools/logs/Get-RecentFolderActivity.ps1 -DriveLetter D -StartTime (Get-Date).AddDays(-7) -Name '案件' -MaxEvents 200

# 日時の範囲を指定してCSV保存（出力先は任意）
./tools/logs/Get-RecentFolderActivity.ps1 -StartTime '2026-11-01 09:00' -EndTime '2026-11-01 18:00' |
  Export-Csv "$env:TEMP\folder-activity.csv" -NoTypeInformation -Encoding UTF8
```

- **対象:** ドライブ文字を持つローカルNTFSボリュームに保存されているUSN変更ジャーナル。共有フォルダ、ネットワークドライブ、ReFS、FAT/exFATは対象外です。標準Windows APIを `tools/logs/lib/UsnJournal.cs` 経由で呼び出します。
- **表示:** `Created`（作成）、`Deleted`（削除）、`RenamedFrom` / `RenamedTo`（名前変更・同一ボリューム内移動の変更前/後）、`MetadataChanged`（属性・時刻変更）、`SecurityChanged`（権限変更）、`ReparsePointChanged`（再解析ポイント変更）。1記録に複数の理由が含まれることがあります。変更前/後は別行で、ボリューム間の移動を1つの移動操作として復元するものではありません。
- **順序と絞り込み:** 日時はローカル時刻。期間は両端を含み、期間・名前を絞った後に `-MaxEvents` を適用し、USNの降順（記録の新しい順）で返します。`-Name` はフォルダ名への部分一致で、ワイルドカードやパス検索ではありません。
- **パス:** 記録された `Name` と `ParentFileId` を保持します。`ParentPathNow` は取得時点の親フォルダのパス、`PathHint` はそこに記録時の名前を結合した参考値です。親や祖先が移動・改名されていると、操作時点のパスと異なります。親が削除済み・参照不可なら両方とも空欄になり、名前とIDだけを表示します。
- **取得限界:** フォルダを開いた・一覧を見た履歴、操作者、プロセス名、コピー元はUSNからは分かりません。USNは監査証跡や完全な操作履歴ではなく、上書き・無効化・削除によって過去の記録がない場合もあります。記録がないことは操作がなかったことを意味しません。
- **読み取りのみ:** ジャーナルの新規作成や有効化、監査ポリシーの変更は行いません。保持されているジャーナル全体を順に読み、実行開始時の末尾までで終了します。保持量が大きい場合は時間がかかり、期間・件数を絞っても読み取り範囲は変わりません。読み取り中に記録が上書きされた場合はエラーになるため再実行してください。
- **エラー時:** 管理者権限、対象ドライブ、既存ジャーナルの有無を確認してください。記録が残っていなければ、このツールでは過去に遡って復元できません。

仕様の根拠: [USN_RECORD_V2](https://learn.microsoft.com/en-us/windows/win32/api/winioctl/ns-winioctl-usn_record_v2)、[変更ジャーナルの読み取り](https://learn.microsoft.com/en-us/windows/win32/fileio/walking-a-buffer-of-change-journal-records)。操作者を調べるには、事前のファイルシステム監査と対象のSACLが必要な[Securityイベント4663](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4663)などを別途調査してください。

## 業務で使うイベントログ

```powershell
# 過去24時間の重大・エラー・警告（全体で最新100件）
./tools/logs/Get-WorkEventLog.ps1 |
  Format-Table TimeCreated, Category, LogName, Id, Message -Wrap

# 再起動とWindows Updateを過去7日間から確認
./tools/logs/Get-WorkEventLog.ps1 -Category Restart,WindowsUpdate -StartTime (Get-Date).AddDays(-7) -MaxEvents 200

# ログオン成功/失敗、ログオフ、ロック/解除（Securityログを読める権限が必要）
./tools/logs/Get-WorkEventLog.ps1 -Category Logon |
  Format-Table TimeCreated, Id, User, LogonType, IpAddress -AutoSize

# 全カテゴリの最新500件をCSV保存
./tools/logs/Get-WorkEventLog.ps1 -Category All -MaxEvents 500 |
  Export-Csv "$env:TEMP\work-events.csv" -NoTypeInformation -Encoding UTF8
```

| `-Category` | 取得する記録 |
| --- | --- |
| `Errors`（既定） | System / Application の重大・エラー・警告（Level 1～3） |
| `Restart` | System: User32 の1074、EventLog の6005/6006/6008、Kernel-Power の41（起動・終了・異常終了などの手掛かり） |
| `Logon` | Security: 4624/4625/4634/4647/4800/4801（成功・失敗・ログオフ・ロック・解除） |
| `WindowsUpdate` | WindowsUpdateClient/Operational の全レベル（更新処理の進行・結果） |
| `Defender` | Windows Defender/Operational の重大・エラー・警告 |
| `All` | 上記すべて |

期間指定はフォルダ履歴と同じです。複数カテゴリの結果を時刻の新しい順に統合し、全体で `-MaxEvents` 件に制限します。同じイベントが複数カテゴリに当たる場合は1行にまとめ、`Category` にカテゴリ名を併記します。通常は一般ユーザーで実行できますが、Securityなどの読み取りには管理者権限やログに対するアクセス権が必要です。取得できないログは警告を出し、残りは処理を続けます。該当記録がない場合の詳細は `-Verbose` で確認できます。

監査が未設定、ログが無効、保持期間外の場合は記録を取得できません。ログオンにはサービスやネットワーク経由のアクセスも含まれるため、勤務時間や対話ログオンの履歴と直接同一視しないでください。メッセージリソースがない記録でも、取得できたID・プロバイダーなどを返します。参考: [再起動のイベントログ調査](https://learn.microsoft.com/en-us/troubleshoot/windows-server/performance/troubleshoot-unexpected-reboots-system-event-logs)、[監査設定とログオンのイベントID](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/advanced-audit-policy-configuration)。

## ログ表示ツールの検証

```powershell
# Pesterが必要（3.4で動作確認）。管理者不要の解析・クエリテスト
Invoke-Pester ./tests/logs/ActivityLogs.Tests.ps1
```

USNの実ドライブ読み取りは管理者として別途確認してください。
