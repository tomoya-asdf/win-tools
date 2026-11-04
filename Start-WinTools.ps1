<#
.SYNOPSIS
番号を選んでWindows用ツールと追加したExcelマクロブックを起動します。
.PARAMETER List
対話せず、現在のメニュー項目をオブジェクトで返します。
.EXAMPLE
.\Start-WinTools.ps1
.EXAMPLE
.\Start-WinTools.ps1 -List | Format-Table Number, Name
#>
[CmdletBinding()]
param([switch]$List)
Import-Module (Join-Path $PSScriptRoot 'tools\launcher\lib\WinToolsMenu.psm1') -Force
if ($List) {
  Get-WinToolsMenuItem -RepositoryRoot $PSScriptRoot
  return
}
while ($true) {
  $items = @(Get-WinToolsMenuItem -RepositoryRoot $PSScriptRoot)
  Write-Host "`n=== Win Tools ==="
  foreach ($item in $items) { Write-Host ('{0,3}: {1}' -f $item.Number, $item.Name) }
  Write-Host '  0: 終了'
  $choice = Read-Host '番号を入力'
  if ($choice -eq '0' -or $null -eq $choice) { return }
  $number = 0
  if (-not [int]::TryParse($choice, [ref]$number)) { Write-Warning '表示された番号を入力してください。'; continue }
  $selected = $items | Where-Object { $_.Number -eq $number } | Select-Object -First 1
  if (-not $selected) { Write-Warning '表示された番号を入力してください。'; continue }
  try { Invoke-WinToolsMenuItem -Item $selected }
  catch { Write-Warning $_.Exception.Message }
  [void](Read-Host 'Enterでメニューへ戻る（Excelマクロの一覧も更新します）')
}
