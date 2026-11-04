<#
.SYNOPSIS
現在のユーザーの右クリックメニューに一時保存を登録します（管理者不要）。
.DESCRIPTION
移動後は再実行して登録先を更新してください。保存先はFolderPathで固定できます。
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$FolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ'))
$toolRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\temporary-save'
Import-Module (Join-Path $toolRoot 'lib\TemporarySave.psm1') -Force
$folder = Get-TemporaryFolderPath -FolderPath $FolderPath
$copyScript = Join-Path $toolRoot 'Copy-TemporaryItem.ps1'
if (-not (Test-Path -LiteralPath $copyScript -PathType Leaf)) { throw "コピー用スクリプトがありません: $copyScript" }
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$command = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}" -LiteralPath "%1" -FolderPath "{2}" -Open' -f $powershell, $copyScript, $folder
foreach ($type in @('*', 'Directory')) {
  $key = "HKCU:\Software\Classes\$type\shell\CopyToSelfFolder"
  if ($PSCmdlet.ShouldProcess($key, 'Register temporary save context menu')) {
    New-Item -Path ($key + '\command') -Force -ErrorAction Stop | Out-Null
    Set-Item -LiteralPath $key -Value 'コピーして開く（一時保存）' -ErrorAction Stop
    Set-Item -LiteralPath ($key + '\command') -Value $command -ErrorAction Stop
  }
}
