function Get-WinToolsMenuItem {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)][string]$RepositoryRoot)
  $definitions = @(
    @{ Name = 'フォルダ比較'; Kind = 'FolderCompare'; Path = 'tools\files\Compare-Folders.ps1' },
    @{ Name = '業務イベントログ'; Kind = 'WorkLogs'; Path = 'tools\logs\Get-WorkEventLog.ps1' },
    @{ Name = 'フォルダ操作履歴（管理者）'; Kind = 'FolderLogs'; Path = 'tools\logs\Get-RecentFolderActivity.ps1' },
    @{ Name = '一時保存へコピー'; Kind = 'TemporaryCopy'; Path = 'tools\temporary-save\Copy-TemporaryItem.ps1' },
    @{ Name = '一時保存のコピー品を削除'; Kind = 'TemporaryClear'; Path = 'tools\temporary-save\Clear-TemporaryFolder.ps1' },
    @{ Name = '一時保存の自動リネーム監視を起動'; Kind = 'TemporaryWatch'; Path = 'tools\temporary-save\Start-TemporaryFolderWatcher.ps1' }
  )
  $number = 1
  foreach ($definition in $definitions) {
    $path = Join-Path $RepositoryRoot $definition.Path
    if (Test-Path -LiteralPath $path -PathType Leaf) {
      [pscustomobject]@{ Number = $number; Name = $definition.Name; Kind = $definition.Kind; Path = $path }
      $number++
    }
  }
  $macroRoot = Join-Path $RepositoryRoot 'tools\excel-macros'
  if (-not (Test-Path -LiteralPath $macroRoot -PathType Container)) { return }
  $pending = [Collections.Generic.Stack[string]]::new()
  $pending.Push($macroRoot)
  $books = [Collections.Generic.List[object]]::new()
  while ($pending.Count -gt 0) {
    $folder = $pending.Pop()
    try { $children = @(Get-ChildItem -LiteralPath $folder -ErrorAction Stop) }
    catch { Write-Warning "Excelマクロの一覧を取得できません: $folder : $($_.Exception.Message)"; continue }
    foreach ($item in $children) {
      if ($item.Name.StartsWith('.') -or $item.Name.StartsWith('~$') -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
      if ($item.Name -like '*.local.*' -or $item.Name -like '*.secrets.*' -or $item.Name -like '*.credentials.*') { continue }
      if ($item.PSIsContainer) {
        if ($item.Name -notin @('data', 'input', 'output', 'local')) { $pending.Push($item.FullName) }
      } elseif ($item.Extension.ToLowerInvariant() -in @('.xlsm', '.xlsb', '.xlam', '.xltm')) {
        $books.Add($item)
      }
    }
  }
  foreach ($book in ($books | Sort-Object FullName)) {
    $relative = $book.FullName.Substring($macroRoot.TrimEnd('\').Length + 1)
    [pscustomobject]@{ Number = $number; Name = 'Excel: ' + $relative; Kind = 'ExcelMacro'; Path = $book.FullName }
    $number++
  }
}

function Read-WinToolsPath {
  param([string]$Prompt)
  $value = ([string](Read-Host $Prompt)).Trim()
  if ($value.Length -ge 2 -and (($value[0] -eq '"' -and $value[$value.Length - 1] -eq '"') -or ($value[0] -eq "'" -and $value[$value.Length - 1] -eq "'"))) {
    $value = $value.Substring(1, $value.Length - 2)
  }
  return $value
}

function Invoke-WinToolsMenuItem {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)]$Item)
  switch ($Item.Kind) {
    'FolderCompare' {
      $left = Read-WinToolsPath '左側（比較元）フォルダ。空欄で戻る'
      if (-not $left) { return }
      $right = Read-WinToolsPath '右側（比較先）フォルダ。空欄で戻る'
      if (-not $right) { return }
      $mode = if ((Read-Host 'SHA-256で内容も比較しますか？ [y/N]') -match '^(y|yes|はい)$') { 'Hash' } else { 'Size' }
      $rows = @(& $Item.Path -LeftPath $left -RightPath $right -CompareBy $mode -DifferencesOnly)
      if ($rows.Count) {
        $rows | Format-Table Status, RelativePath, LeftLength, RightLength, Note -AutoSize -Wrap | Out-Host
      } else {
        Write-Host '指定した比較方法で差分は見つかりませんでした。'
      }
      if ($mode -eq 'Size') { Write-Host '同じサイズのファイルについて、内容の一致は確認していません。' }
    }
    'WorkLogs' {
      $category = Read-Host 'カテゴリ [Errors / Restart / Logon / WindowsUpdate / Defender / All]（空欄: Errors）'
      if (-not $category) { $category = 'Errors' }
      & $Item.Path -Category $category | Format-Table TimeCreated, Category, Id, Message -Wrap | Out-Host
    }
    'FolderLogs' {
      & $Item.Path | Format-Table TimeCreated, Operation, Name, PathHint -AutoSize -Wrap | Out-Host
    }
    'TemporaryCopy' {
      $source = Read-WinToolsPath 'コピーするファイル・フォルダ。空欄で戻る'
      if (-not $source) { return }
      $folder = Read-WinToolsPath '保存先（空欄: ユーザープロファイルの一時保存フォルダ）'
      $arguments = @{ LiteralPath = $source; Open = $true }
      if ($folder) { $arguments.FolderPath = $folder }
      & $Item.Path @arguments | Format-Table Name, FullName -AutoSize -Wrap | Out-Host
    }
    'TemporaryClear' {
      $folder = Read-WinToolsPath '保存先（空欄: ユーザープロファイルの一時保存フォルダ）'
      $arguments = @{}
      if ($folder) { $arguments.FolderPath = $folder }
      & $Item.Path @arguments
    }
    'TemporaryWatch' {
      $folder = Read-WinToolsPath '監視先（空欄: ユーザープロファイルの一時保存フォルダ）'
      $helper = Join-Path (Split-Path -Parent $Item.Path) 'lib\TemporarySave.psm1'
      Import-Module $helper -Force
      $folder = if ($folder) { Get-TemporaryFolderPath -FolderPath $folder } else { Get-TemporaryFolderPath }
      $arguments = @('-NoProfile', '-Sta', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $Item.Path), '-FolderPath', ('"{0}"' -f $folder))
      Start-Process -FilePath powershell.exe -ArgumentList $arguments -WindowStyle Hidden -ErrorAction Stop
      Write-Host '自動リネーム監視を起動しました。終了するにはトレイアイコンのExitを選択してください。'
    }
    'ExcelMacro' {
      $file = Get-Item -LiteralPath $Item.Path -Force -ErrorAction Stop
      if ($file.PSIsContainer -or $file.Extension.ToLowerInvariant() -notin @('.xlsm', '.xlsb', '.xlam', '.xltm')) { throw '対応するExcelマクロファイルを指定してください。' }
      Start-Process -FilePath $file.FullName -ErrorAction Stop
    }
    default { throw "未対応のメニュー項目です: $($Item.Kind)" }
  }
}

Export-ModuleMember -Function Get-WinToolsMenuItem, Invoke-WinToolsMenuItem
