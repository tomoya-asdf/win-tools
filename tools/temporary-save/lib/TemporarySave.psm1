function Get-TemporaryFolderPath {
  param([string]$FolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ'))
  $provider = $null
  $drive = $null
  $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FolderPath, [ref]$provider, [ref]$drive)
  if ($provider.Name -ne 'FileSystem') { throw '保存先にはファイルシステムのパスを指定してください。' }
  $path = [IO.Path]::GetFullPath($path).TrimEnd('\')
  if ($path -eq [IO.Path]::GetPathRoot($path).TrimEnd('\')) { throw 'ドライブや共有のルートを保存先には指定できません。' }
  return $path
}

function Get-TemporaryItemName {
  param([Parameter(Mandatory=$true)][string]$Name)
  return ('【一時保存】＜{0}＞{1}' -f (Get-Date -Format yyyyMMdd), $Name)
}

function Rename-TemporaryItem {
  param([Parameter(Mandatory=$true)][string]$LiteralPath)
  if (-not (Test-Path -LiteralPath $LiteralPath)) { return }
  $item = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
  if ($item.Name -match '^(~|\d{3}_|【一時保存】)') { return }
  $newName = Get-TemporaryItemName -Name $item.Name
  if (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $item.FullName) $newName) -ErrorAction SilentlyContinue) {
    throw "同名の一時保存項目が既に存在します: $newName"
  }
  Rename-Item -LiteralPath $item.FullName -NewName $newName -ErrorAction Stop
}

Export-ModuleMember -Function Get-TemporaryFolderPath, Get-TemporaryItemName, Rename-TemporaryItem
