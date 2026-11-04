<#
.SYNOPSIS
一時保存先に新規作成された項目を日付付きの名前に変更します。
.DESCRIPTION
既定はトレイ常駐（Exitで終了）。NoTrayはCtrl+Cで終了します。
既存項目とサブフォルダ内は処理しません。保存先はコードの場所と独立しています。
.EXAMPLE
.\tools\temporary-save\Start-TemporaryFolderWatcher.ps1 -NoTray -DurationSeconds 10
#>
[CmdletBinding()]
param(
  [string]$FolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ'),
  [switch]$NoTray,
  [ValidateRange(0, 2147483647)][int]$DurationSeconds = 0
)
Import-Module (Join-Path $PSScriptRoot 'lib\TemporarySave.psm1') -Force
$folder = Get-TemporaryFolderPath -FolderPath $FolderPath
New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
$watcher = New-Object IO.FileSystemWatcher
$watcher.Path = $folder
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [IO.NotifyFilters]'FileName, DirectoryName'
$sourceId = 'WinTools.TemporarySave.' + [guid]::NewGuid().ToString('N')
$errorId = $sourceId + '.Error'
$pending = New-Object Collections.ArrayList
$started = Get-Date
$timer = $null
$icon = $null
$menu = $null
$context = $null

function Receive-TemporaryFolderEvents {
  foreach ($event in @(Get-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue)) {
    [void]$pending.Add(@{ Path = $event.SourceEventArgs.FullPath; Due = (Get-Date).AddSeconds(1); Attempts = 0 })
    Remove-Event -EventIdentifier $event.EventIdentifier
  }
  foreach ($event in @(Get-Event -SourceIdentifier $errorId -ErrorAction SilentlyContinue)) {
    Write-Warning "監視中にイベントを取得できませんでした: $($event.SourceEventArgs.GetException().Message)"
    Remove-Event -EventIdentifier $event.EventIdentifier
  }
  foreach ($entry in @($pending.ToArray())) {
    if ((Get-Date) -lt $entry.Due) { continue }
    try {
      Rename-TemporaryItem -LiteralPath $entry.Path
      $pending.Remove($entry)
    } catch {
      $entry.Attempts++
      $entry.Due = (Get-Date).AddSeconds(1)
      if ($entry.Attempts -ge 5) {
        Write-Warning "名前を変更できません: $($entry.Path): $($_.Exception.Message)"
        $pending.Remove($entry)
      }
    }
  }
}

try {
  Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier $sourceId | Out-Null
  Register-ObjectEvent -InputObject $watcher -EventName Error -SourceIdentifier $errorId | Out-Null
  $watcher.EnableRaisingEvents = $true
  Write-Verbose "監視先: $folder"
  if ($NoTray) {
    while ($DurationSeconds -eq 0 -or ((Get-Date) - $started).TotalSeconds -lt $DurationSeconds) {
      Receive-TemporaryFolderEvents
      Start-Sleep -Milliseconds 250
    }
  } else {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $context = New-Object Windows.Forms.ApplicationContext
    $menu = New-Object Windows.Forms.ContextMenuStrip
    $exitItem = $menu.Items.Add('Exit')
    $exitItem.add_Click({ $context.ExitThread() })
    $icon = New-Object Windows.Forms.NotifyIcon
    $icon.Icon = [Drawing.SystemIcons]::Application
    $icon.Text = 'WinTools Temporary Save'
    $icon.ContextMenuStrip = $menu
    $icon.Visible = $true
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 250
    $timer.add_Tick({
      Receive-TemporaryFolderEvents
      if ($DurationSeconds -gt 0 -and ((Get-Date) - $started).TotalSeconds -ge $DurationSeconds) { $context.ExitThread() }
    })
    $timer.Start()
    [Windows.Forms.Application]::Run($context)
  }
} finally {
  $watcher.EnableRaisingEvents = $false
  Unregister-Event -SourceIdentifier $sourceId -ErrorAction SilentlyContinue
  Unregister-Event -SourceIdentifier $errorId -ErrorAction SilentlyContinue
  Get-Event -ErrorAction SilentlyContinue | Where-Object { $_.SourceIdentifier -eq $sourceId -or $_.SourceIdentifier -eq $errorId } | Remove-Event
  $watcher.Dispose()
  if ($timer) { $timer.Stop(); $timer.Dispose() }
  if ($icon) { $icon.Visible = $false; $icon.Dispose() }
  if ($menu) { $menu.Dispose() }
  if ($context) { $context.Dispose() }
}
