<#
.SYNOPSIS
番号を選んでWindows用ツール・追加アプリ・Excelマクロを起動します。
.PARAMETER List
対話せず、現在のメニュー項目をオブジェクトで返します。
.EXAMPLE
.\Start-WinTools.ps1
.EXAMPLE
.\Start-WinTools.ps1 -List | Format-Table Number, Name
#>
[CmdletBinding()]
param([switch]$List)
Import-Module (Join-Path $PSScriptRoot 'tools\launcher\lib\WinToolsMenu.psm1')
if ($List) {
  Get-WinToolsMenuItem -RepositoryRoot $PSScriptRoot
  return
}
$appRoot = Join-Path $PSScriptRoot 'apps'
$validationMessage = $null
try { $null = [IO.Directory]::CreateDirectory($appRoot) }
catch { $validationMessage = 'アプリ登録フォルダを作成できません: ' + $_.Exception.Message }
$appFolder = $null
while ($true) {
  Clear-Host
  if ($appFolder) {
    try { $items = @(Get-WinToolsAppMenuItem -Directory $appFolder -Root $appRoot) }
    catch {
      $appFolder = $null
      $validationMessage = $_.Exception.Message
      continue
    }
    $relative = $appFolder.Substring($appRoot.Length).TrimStart('\')
    Write-Host ("`n=== Win Tools / アプリ / {0} ===" -f $relative)
    if (-not $items.Count) { Write-Host 'このフォルダには項目がありません。' }
  } else {
    $items = @(Get-WinToolsMenuItem -RepositoryRoot $PSScriptRoot)
    Write-Host "`n=== Win Tools ==="
  }
  foreach ($item in $items) { Write-Host ('{0,3}: {1}' -f $item.Number, $item.Name) }
  if ($appFolder) { Write-Host '  0: 一つ上へ戻る' } else { Write-Host '  0: 終了' }
  if ($validationMessage) { Write-Warning $validationMessage; $validationMessage = $null }
  $choice = Read-Host '番号を入力'
  if ($null -eq $choice) { return }
  if ($choice -eq '0') {
    if (-not $appFolder) { return }
    $parent = Split-Path -Parent $appFolder
    if ($parent.TrimEnd('\') -eq $appRoot.TrimEnd('\')) { $appFolder = $null } else { $appFolder = $parent }
    continue
  }
  $number = 0
  if (-not [int]::TryParse($choice, [ref]$number)) { $validationMessage = '表示された番号を入力してください。'; continue }
  $selected = $items | Where-Object { $_.Number -eq $number } | Select-Object -First 1
  if (-not $selected) { $validationMessage = '表示された番号を入力してください。'; continue }
  if ($selected.Kind -eq 'AppFolder') { $appFolder = $selected.Path; continue }
  try { Invoke-WinToolsMenuItem -Item $selected }
  catch { Write-Warning $_.Exception.Message }
  [void](Read-Host 'Enterでメニューへ戻る')
}
