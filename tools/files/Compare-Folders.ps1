<#
.SYNOPSIS
2つのフォルダを相対パス・種類・サイズ、必要に応じてSHA-256で比較します。
.DESCRIPTION
隠し項目と空フォルダも対象です。子のJunctionやSymbolicLinkはたどりません。
SameSizeは内容の一致を意味しません。内容の確認にはCompareBy Hashを指定します。
列挙・読み取りできない場所は結果に残し、その配下を片側だけの項目とは断定しません。
.EXAMPLE
.\tools\files\Compare-Folders.ps1 -LeftPath 'C:\Work\original' -RightPath 'D:\Backup' -CompareBy Hash -DifferencesOnly
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$LeftPath,
  [Parameter(Mandatory=$true)][string]$RightPath,
  [ValidateSet('Size', 'Hash')][string]$CompareBy = 'Size',
  [switch]$DifferencesOnly
)

function Get-ComparisonRoot {
  param([string]$Path)
  $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
  if ($item.PSProvider.Name -ne 'FileSystem' -or -not $item.PSIsContainer) {
    throw "ファイルシステムのフォルダを指定してください: $Path"
  }
  return $item.FullName
}

function Get-FolderInventory {
  param([string]$Root)
  $entries = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
  $blocked = [Collections.Generic.List[string]]::new()
  $pending = [Collections.Generic.Stack[string]]::new()
  $pending.Push('')
  while ($pending.Count -gt 0) {
    $relative = $pending.Pop()
    $path = if ($relative) { Join-Path $Root $relative } else { $Root }
    try {
      $children = @(Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop)
    } catch {
      $key = if ($relative) { $relative } else { '.' }
      $blocked.Add($key)
      if (-not $entries.ContainsKey($key)) {
        $entries[$key] = [pscustomobject]@{ Kind = 'Directory'; FullName = $path; Length = $null; LastWriteTimeUtc = $null; Error = $null }
      }
      $entries[$key].Error = $_.Exception.Message
      continue
    }
    foreach ($child in $children) {
      $key = if ($relative) { $relative + '\' + $child.Name } else { $child.Name }
      $isLink = ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
      $kind = if ($isLink) { 'Link' } elseif ($child.PSIsContainer) { 'Directory' } else { 'File' }
      $entries[$key] = [pscustomobject]@{
        Kind = $kind
        FullName = $child.FullName
        Length = if ($kind -eq 'File') { $child.Length } else { $null }
        LastWriteTimeUtc = $child.LastWriteTimeUtc
        Error = $null
      }
      if ($isLink) { $blocked.Add($key) }
      elseif ($child.PSIsContainer) { $pending.Push($key) }
    }
  }
  return [pscustomobject]@{ Entries = $entries; Blocked = $blocked }
}

function Test-BlockedAncestor {
  param([string]$RelativePath, $Inventory)
  foreach ($path in $Inventory.Blocked) {
    if ($path -eq '.' -or $RelativePath.StartsWith($path + '\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
  }
  return $false
}

function Get-StableFileHash {
  param($Entry)
  $before = Get-Item -LiteralPath $Entry.FullName -Force -ErrorAction Stop
  if ($before.Length -ne $Entry.Length -or $before.LastWriteTimeUtc -ne $Entry.LastWriteTimeUtc) {
    throw '列挙後にファイルが変更されました。再度比較してください。'
  }
  $hash = (Get-FileHash -LiteralPath $Entry.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
  $after = Get-Item -LiteralPath $Entry.FullName -Force -ErrorAction Stop
  if (-not $hash -or $after.Length -ne $before.Length -or $after.LastWriteTimeUtc -ne $before.LastWriteTimeUtc) {
    throw '読み取り中にファイルが変更されたか、ハッシュを取得できませんでした。'
  }
  return $hash
}

$leftRoot = Get-ComparisonRoot -Path $LeftPath
$rightRoot = Get-ComparisonRoot -Path $RightPath
$left = Get-FolderInventory -Root $leftRoot
$right = Get-FolderInventory -Root $rightRoot
$paths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($path in $left.Entries.Keys) { [void]$paths.Add($path) }
foreach ($path in $right.Entries.Keys) { [void]$paths.Add($path) }

foreach ($path in ($paths | Sort-Object)) {
  $leftItem = $null
  $rightItem = $null
  [void]$left.Entries.TryGetValue($path, [ref]$leftItem)
  [void]$right.Entries.TryGetValue($path, [ref]$rightItem)
  $note = ''
  $leftHash = $null
  $rightHash = $null
  if (($leftItem -and $leftItem.Error) -or ($rightItem -and $rightItem.Error)) {
    $status = 'ReadError'
    $note = (@($leftItem.Error, $rightItem.Error) | Where-Object { $_ }) -join ' / '
  } elseif ((Test-BlockedAncestor -RelativePath $path -Inventory $left) -or (Test-BlockedAncestor -RelativePath $path -Inventory $right)) {
    $status = 'NotCompared'
    $note = '親フォルダの列挙不可、または親がリンクのため比較できません。'
  } elseif (($leftItem -and $leftItem.Kind -eq 'Link') -or ($rightItem -and $rightItem.Kind -eq 'Link')) {
    $status = 'SkippedLink'
    $note = 'Junction・SymbolicLinkなどの再解析ポイントはたどりません。'
  } elseif (-not $leftItem) {
    $status = 'OnlyRight'
  } elseif (-not $rightItem) {
    $status = 'OnlyLeft'
  } elseif ($leftItem.Kind -ne $rightItem.Kind) {
    $status = 'TypeMismatch'
  } elseif ($leftItem.Kind -eq 'Directory') {
    $status = 'Directory'
  } elseif ($leftItem.Length -ne $rightItem.Length) {
    $status = 'DifferentSize'
  } elseif ($CompareBy -eq 'Size') {
    $status = 'SameSize'
  } else {
    try {
      $leftHash = Get-StableFileHash -Entry $leftItem
      $rightHash = Get-StableFileHash -Entry $rightItem
      $status = if ($leftHash -eq $rightHash) { 'SameContent' } else { 'DifferentContent' }
    } catch {
      $status = 'ReadError'
      $note = $_.Exception.Message
    }
  }
  if ($DifferencesOnly -and $status -in @('Directory', 'SameSize', 'SameContent')) { continue }
  [pscustomobject]@{
    RelativePath = $path
    Status = $status
    LeftKind = $leftItem.Kind
    RightKind = $rightItem.Kind
    LeftLength = $leftItem.Length
    RightLength = $rightItem.Length
    LeftLastWriteTime = if ($leftItem.LastWriteTimeUtc) { $leftItem.LastWriteTimeUtc.ToLocalTime() } else { $null }
    RightLastWriteTime = if ($rightItem.LastWriteTimeUtc) { $rightItem.LastWriteTimeUtc.ToLocalTime() } else { $null }
    LeftHash = $leftHash
    RightHash = $rightHash
    Note = $note
  }
}
