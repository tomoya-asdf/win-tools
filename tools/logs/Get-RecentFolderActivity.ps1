<#
.SYNOPSIS
既存のNTFS USNジャーナルから直近のフォルダ変更を表示します（管理者用）。
.DESCRIPTION
作成・削除・名前変更・属性/権限変更をUSNの降順で返します。既定は過去24時間、100件。
フォルダを開いた履歴、操作者、プロセス名はUSNには記録されません。
PathHintは現在の親パスと記録された名前から作る参考値で、操作時点の完全パスではありません。
ジャーナルや監査の設定は変更しません。保持済みの履歴全体を読むため、大きなジャーナルでは時間がかかります。
.PARAMETER DriveLetter
対象のローカルNTFSドライブ。既定はシステムドライブ。
.PARAMETER StartTime
取得期間の開始。既定は24時間前。
.PARAMETER EndTime
取得期間の終了。既定は実行時刻。
.PARAMETER MaxEvents
名前と期間で絞り込んだ後に返す最大件数。
.PARAMETER Name
記録されたフォルダ名の部分一致（大文字小文字を区別しないリテラル文字列）。
.EXAMPLE
.\tools\logs\Get-RecentFolderActivity.ps1 | Format-Table TimeCreated,Operation,Name,PathHint -AutoSize -Wrap
.EXAMPLE
.\tools\logs\Get-RecentFolderActivity.ps1 -DriveLetter D -StartTime (Get-Date).AddDays(-7) -Name '案件'
#>
[CmdletBinding()]
param(
  [ValidatePattern('^[A-Za-z]:?$')]
  [string]$DriveLetter = $env:SystemDrive,
  [datetime]$StartTime = (Get-Date).AddHours(-24),
  [datetime]$EndTime = (Get-Date),
  [ValidateRange(1, 100000)]
  [int]$MaxEvents = 100,
  [string]$Name
)

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
  throw 'このスクリプトはWindows専用です。'
}
if ($StartTime -gt $EndTime) { throw 'StartTimeはEndTime以前を指定してください。' }
$drive = $DriveLetter.TrimEnd(':').ToUpperInvariant() + ':'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'USNジャーナルの読み取りには管理者権限が必要です。PowerShellを「管理者として実行」して再実行してください。'
  }
} finally {
  $identity.Dispose()
}
$volume = New-Object System.IO.DriveInfo($drive + '\')
if (-not $volume.IsReady -or $volume.DriveFormat -ne 'NTFS' -or $volume.DriveType -eq 'Network') {
  throw "$drive は読み取り可能なローカルNTFSドライブではありません。"
}
if (-not ('WinTools.UsnJournal' -as [type])) {
  Add-Type -Path (Join-Path $PSScriptRoot 'lib\UsnJournal.cs') -ErrorAction Stop
}
Write-Verbose "$drive の既存USNジャーナルを読み取ります。期間: $StartTime ～ $EndTime"
try {
  $changes = [WinTools.UsnJournal]::Read($drive, $StartTime, $EndTime, $MaxEvents, $Name)
} catch {
  throw "USN履歴を取得できませんでした。ジャーナルが無効・削除済み、アクセス拒否などが考えられます。詳細: $($_.Exception.GetBaseException().Message)"
}
if ($changes.Count -eq 0) {
  Write-Warning '指定条件のフォルダ変更は保存済みUSNジャーナルに見つかりませんでした。古い記録が上書きされた可能性もあります。'
}
$changes | Select-Object TimeCreated, Operation, Name, PathHint, ParentPathNow, FileId, ParentFileId, Usn, ReasonMask
