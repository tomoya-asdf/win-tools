<#
.SYNOPSIS
一時保存先の直下にある日付付きのコピー品を、確認して削除します。
.EXAMPLE
.\tools\temporary-save\Clear-TemporaryFolder.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param([string]$FolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ'))
Import-Module (Join-Path $PSScriptRoot 'lib\TemporarySave.psm1') -Force
$folder = Get-TemporaryFolderPath -FolderPath $FolderPath
if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return }

function Assert-NoReparsePoint {
  param([IO.FileSystemInfo]$Item)
  if (($Item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
    throw "リンクを含む項目は自動削除できません。内容を確認してください: $($Item.FullName)"
  }
  if ($Item -is [IO.DirectoryInfo]) {
    foreach ($child in $Item.GetFileSystemInfos()) { Assert-NoReparsePoint -Item $child }
  }
}

foreach ($item in (Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop)) {
  if ($item.Name -notmatch '^【一時保存】＜\d{8}＞') { continue }
  # Inspect without following junctions before removing a tree.
  Assert-NoReparsePoint -Item $item
  if ($PSCmdlet.ShouldProcess($item.FullName, 'Delete temporary saved item')) {
    Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
  }
}
