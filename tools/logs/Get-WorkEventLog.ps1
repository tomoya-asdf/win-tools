<#
.SYNOPSIS
障害調査や作業確認に使うWindowsイベントログを表示します。
.DESCRIPTION
既定は過去24時間のSystem/Applicationの重大・エラー・警告を合計100件。
オブジェクトを返すのでFormat-Table、Export-Csvなどに渡せます。
取得できないログは警告にし、取得可能なカテゴリの処理を継続します。
.PARAMETER Category
Errors、Restart、Logon、WindowsUpdate、Defender、All。複数指定できます。
LogonはSecurityログの読み取り権限と事前の監査設定が必要です。
.PARAMETER StartTime
取得期間の開始。既定は24時間前。
.PARAMETER EndTime
取得期間の終了。既定は実行時刻。
.PARAMETER MaxEvents
全カテゴリを合わせた最大件数（新しい順）。
.EXAMPLE
.\tools\logs\Get-WorkEventLog.ps1 | Format-Table TimeCreated,Category,Id,Message -Wrap
.EXAMPLE
.\tools\logs\Get-WorkEventLog.ps1 -Category Restart,WindowsUpdate -StartTime (Get-Date).AddDays(-7)
.EXAMPLE
.\tools\logs\Get-WorkEventLog.ps1 -Category Logon | Select-Object TimeCreated,Id,User,LogonType,IpAddress
#>
[CmdletBinding()]
param(
  [ValidateSet('Errors', 'Restart', 'Logon', 'WindowsUpdate', 'Defender', 'All')]
  [string[]]$Category = @('Errors'),
  [datetime]$StartTime = (Get-Date).AddHours(-24),
  [datetime]$EndTime = (Get-Date),
  [ValidateRange(1, 100000)]
  [int]$MaxEvents = 100
)

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
  throw 'このスクリプトはWindows専用です。'
}
if ($StartTime -gt $EndTime) { throw 'StartTimeはEndTime以前を指定してください。' }
# Provider names prevent unrelated providers with the same event ID from matching.
$queries = @{
  Errors = @(
    @{ LogName = 'System'; Level = @(1, 2, 3) },
    @{ LogName = 'Application'; Level = @(1, 2, 3) }
  )
  Restart = @(
    @{ LogName = 'System'; ProviderName = 'User32'; Id = @(1074) },
    @{ LogName = 'System'; ProviderName = 'EventLog'; Id = @(6005, 6006, 6008) },
    @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = @(41) }
  )
  Logon = @(
    @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = @(4624, 4625, 4634, 4647, 4800, 4801) }
  )
  WindowsUpdate = @(
    @{ LogName = 'Microsoft-Windows-WindowsUpdateClient/Operational' }
  )
  Defender = @(
    @{ LogName = 'Microsoft-Windows-Windows Defender/Operational'; Level = @(1, 2, 3) }
  )
}
$selected = if ($Category -contains 'All') {
  @('Errors', 'Restart', 'Logon', 'WindowsUpdate', 'Defender')
} else { @($Category | Select-Object -Unique) }

$results = foreach ($kind in $selected) {
  foreach ($query in $queries[$kind]) {
    $filter = $query.Clone()
    $filter.StartTime = $StartTime
    $filter.EndTime = $EndTime
    try {
      # Limit each query, then sort/limit globally so a busy log cannot hide others.
      $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEvents -ErrorAction Stop)
    } catch {
      if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') {
        Write-Verbose "$kind : 指定期間に該当イベントがありません。"
      } else {
        Write-Warning ("{0} ({1}) を取得できません。権限・ログの有効状態を確認してください。詳細: {2}" -f $kind, $query.LogName, $_.Exception.Message)
      }
      continue
    }
    foreach ($event in $events) {
      $data = @{}
      try {
        $xml = [xml]$event.ToXml()
        foreach ($node in $xml.SelectNodes('/*[local-name()="Event"]/*[local-name()="EventData"]/*[local-name()="Data"][@Name]')) {
          $data[$node.GetAttribute('Name')] = $node.InnerText
        }
      } catch {
        Write-Verbose "イベント $($event.RecordId) の補助データを取得できません: $($_.Exception.Message)"
      }
      $user = $data['TargetUserName']
      $domain = $data['TargetDomainName']
      if (-not $user -or $user -eq '-') {
        $user = $data['SubjectUserName']
        $domain = $data['SubjectDomainName']
      }
      if ($user -and $domain -and $domain -ne '-') { $user = "$domain\$user" }
      $message = $null
      try { $message = $event.Message } catch { Write-Verbose 'メッセージリソースがありません。' }
      [pscustomobject]@{
        TimeCreated = $event.TimeCreated
        Category = $kind
        LogName = $event.LogName
        ProviderName = $event.ProviderName
        Id = $event.Id
        Level = $event.LevelDisplayName
        User = $user
        LogonType = $data['LogonType']
        IpAddress = $data['IpAddress']
        Message = $message
        RecordId = $event.RecordId
        MachineName = $event.MachineName
      }
    }
  }
}
# The same event may match Errors and Restart; preserve the category labels together.
$merged = foreach ($group in ($results | Group-Object -Property MachineName, LogName, RecordId)) {
  $row = $group.Group[0]
  $row.Category = ($group.Group.Category | Select-Object -Unique) -join ', '
  $row
}
$merged | Sort-Object TimeCreated, RecordId -Descending | Select-Object -First $MaxEvents
