<#
.SYNOPSIS
Win Toolsの実行履歴を新しい順に返します。既定は直近50件です。
.DESCRIPTION
ログはユーザー別のLOCALAPPDATA\WinTools\Historyに月別のJSON Linesで保存します。
Excelや常駐監視の履歴は起動結果であり、別プロセスの処理完了を示すものではありません。
#>
[CmdletBinding()]
param(
  [ValidateRange(1, 10000)][int]$Last = 50,
  [string]$Directory = (Join-Path $env:LOCALAPPDATA 'WinTools\History')
)
Import-Module (Join-Path $PSScriptRoot 'lib\WinToolsHistory.psm1')
Get-WinToolsHistory -Last $Last -Directory $Directory
