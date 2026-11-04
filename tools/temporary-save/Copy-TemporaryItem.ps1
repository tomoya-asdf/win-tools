<#
.SYNOPSIS
ファイルやフォルダを、日付付きの名前で一時保存先にコピーします。
.EXAMPLE
.\tools\temporary-save\Copy-TemporaryItem.ps1 -LiteralPath 'C:\Work\案件.xlsx'
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [Parameter(Mandatory=$true, Position=0)][string]$LiteralPath,
  [string]$FolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ'),
  [switch]$Open
)
Import-Module (Join-Path $PSScriptRoot 'lib\TemporarySave.psm1') -Force
$folder = Get-TemporaryFolderPath -FolderPath $FolderPath
$source = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
$sourcePath = $source.FullName.TrimEnd('\')
if ($source.PSIsContainer -and ($folder -eq $sourcePath -or $folder.StartsWith($sourcePath + '\', [StringComparison]::OrdinalIgnoreCase))) {
  throw 'コピー元自身やその配下を一時保存先には指定できません。'
}
$destination = Join-Path $folder (Get-TemporaryItemName -Name $source.Name)
if (Test-Path -LiteralPath $destination) { throw "保存先に同名の項目が既に存在します: $destination" }
if ($PSCmdlet.ShouldProcess($destination, "Copy from '$($source.FullName)'")) {
  New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
  Copy-Item -LiteralPath $source.FullName -Destination $destination -Recurse -ErrorAction Stop
  if ($Open) {
    Start-Process -FilePath explorer.exe -ArgumentList ('/select,"{0}"' -f $destination) -ErrorAction Stop
  }
  Get-Item -LiteralPath $destination -Force
}
