function Get-WinToolsMenuItem {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)][string]$RepositoryRoot)
  $definitions = @(
    @{ Name = 'フォルダ比較'; Kind = 'FolderCompare'; Path = 'tools\files\Compare-Folders.ps1' },
    @{ Name = '業務イベントログ'; Kind = 'WorkLogs'; Path = 'tools\logs\Get-WorkEventLog.ps1' },
    @{ Name = 'フォルダ操作履歴（管理者）'; Kind = 'FolderLogs'; Path = 'tools\logs\Get-RecentFolderActivity.ps1' },
    @{ Name = '一時保存へコピー'; Kind = 'TemporaryCopy'; Path = 'tools\temporary-save\Copy-TemporaryItem.ps1' },
    @{ Name = '一時保存のコピー品を削除'; Kind = 'TemporaryClear'; Path = 'tools\temporary-save\Clear-TemporaryFolder.ps1' },
    @{ Name = '一時保存の自動リネーム監視を起動'; Kind = 'TemporaryWatch'; Path = 'tools\temporary-save\Start-TemporaryFolderWatcher.ps1' },
    @{ Name = 'Win Toolsの実行履歴'; Kind = 'ExecutionHistory'; Path = 'tools\launcher\Get-WinToolsHistory.ps1' }
  )
  $number = 1
  foreach ($definition in $definitions) {
    $path = Join-Path $RepositoryRoot $definition.Path
    if (Test-Path -LiteralPath $path -PathType Leaf) {
      [pscustomobject]@{ Number = $number; Name = $definition.Name; Kind = $definition.Kind; Path = $path }
      $number++
    }
  }
  $appRoot = Join-Path $RepositoryRoot 'apps'
  if (Test-Path -LiteralPath $appRoot -PathType Container) {
    try {
      foreach ($item in @(Get-WinToolsAppMenuItem -Directory $appRoot -Root $appRoot)) {
        $item.Number = $number
        $item
        $number++
      }
    } catch { Write-Warning ('追加アプリの一覧を取得できません: ' + $_.Exception.Message) }
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

Import-Module (Join-Path $PSScriptRoot 'WinToolsHistory.psm1')

function Get-WinToolsAppPath {
  [CmdletBinding()]
  param([string]$Path, [string]$Root)
  $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
  $fullPath = [IO.Path]::GetFullPath($Path).TrimEnd('\')
  if (-not $fullPath.Equals($rootPath, [StringComparison]::OrdinalIgnoreCase) -and
      -not $fullPath.StartsWith($rootPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw '追加アプリのフォルダ外は選択できません。'
  }
  $item = Get-Item -LiteralPath $fullPath -Force -ErrorAction Stop
  $current = $item
  while ($null -ne $current) {
    if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw '追加アプリのJunction・SymbolicLinkはたどりません。ショートカット（.lnk）を利用してください。'
    }
    if ($current.FullName.TrimEnd('\').Equals($rootPath, [StringComparison]::OrdinalIgnoreCase)) { return $item }
    if ($current -is [IO.DirectoryInfo]) { $current = $current.Parent } else { $current = $current.Directory }
  }
  throw '追加アプリの登録先を確認できません。'
}

function Get-WinToolsAppMenuItem {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Directory,
    [string]$Root = $Directory
  )
  $folder = Get-WinToolsAppPath -Path $Directory -Root $Root
  if (-not $folder.PSIsContainer) { throw 'アプリ一覧にはフォルダを指定してください。' }
  $children = @(Get-ChildItem -LiteralPath $folder.FullName -ErrorAction Stop |
    Where-Object { -not $_.Name.StartsWith('.') -and ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 } |
    Sort-Object @{ Expression = { -not $_.PSIsContainer } }, Name)
  $number = 1
  foreach ($child in $children) {
    $kind = if ($child.PSIsContainer) { 'AppFolder' } else { 'UserApp' }
    $label = if ($child.PSIsContainer) { 'フォルダ: ' } else { 'アプリ: ' }
    [pscustomobject]@{
      Number = $number
      Name = $label + $child.Name
      Kind = $kind
      Path = $child.FullName
      AppRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    }
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
  $started = [DateTimeOffset]::Now
  $timer = [Diagnostics.Stopwatch]::StartNew()
  $outcome = 'Cancelled'
  $message = $null
  try {
    Clear-Host
    Write-Host ('=== {0} ===' -f $Item.Name)
    Invoke-WinToolsMenuAction -Item $Item -Outcome ([ref]$outcome)
  } catch {
    $outcome = 'Failed'
    $message = $_.Exception.Message
    throw
  } finally {
    $timer.Stop()
    $record = [pscustomobject]@{
      SchemaVersion = 1
      Id = [guid]::NewGuid().ToString('N')
      StartTime = $started.ToString('o')
      EndTime = [DateTimeOffset]::Now.ToString('o')
      Name = $Item.Name
      Kind = $Item.Kind
      ToolPath = $Item.Path
      Status = $outcome
      DurationMs = $timer.ElapsedMilliseconds
      Error = $message
    }
    try { Write-WinToolsHistory -Record $record -ErrorAction Stop }
    catch { Write-Warning ('実行履歴を保存できませんでした: ' + $_.Exception.Message) -WarningAction Continue }
  }
}

function Invoke-WinToolsMenuAction {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)]$Item, [Parameter(Mandatory=$true)][ref]$Outcome)
  $ErrorActionPreference = 'Stop'
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
      $Outcome.Value = 'Launched'
      Write-Host '自動リネーム監視を起動しました。終了するにはトレイアイコンのExitを選択してください。'
    }
    'ExcelMacro' {
      $file = Get-Item -LiteralPath $Item.Path -Force -ErrorAction Stop
      if ($file.PSIsContainer -or $file.Extension.ToLowerInvariant() -notin @('.xlsm', '.xlsb', '.xlam', '.xltm')) { throw '対応するExcelマクロファイルを指定してください。' }
      if ($file.Extension -eq '.xltm') {
        Start-Process -FilePath $file.FullName -Verb New -ErrorAction Stop
      } else {
        Start-Process -FilePath $file.FullName -ErrorAction Stop
      }
    }
    'UserApp' {
      $file = Get-WinToolsAppPath -Path $Item.Path -Root $Item.AppRoot
      if ($file.PSIsContainer) { throw 'フォルダは子メニューから選択してください。' }
      Start-Process -FilePath $file.FullName -WorkingDirectory $file.DirectoryName -ErrorAction Stop
      $Outcome.Value = 'Launched'
    }
    'ExecutionHistory' {
      $rows = @(& $Item.Path)
      if ($rows.Count) {
        $rows | Select-Object @{ Name = '実行日時'; Expression = { ([DateTimeOffset]::Parse($_.StartTime)).LocalDateTime } },
          Name, Status, DurationMs, Error | Format-Table -AutoSize -Wrap | Out-Host
      } else {
        Write-Host '実行履歴はまだありません。'
      }
    }
    default { throw "未対応のメニュー項目です: $($Item.Kind)" }
  }
  if ($Item.Kind -eq 'ExcelMacro') { $Outcome.Value = 'Launched' }
  if ($Outcome.Value -eq 'Cancelled') { $Outcome.Value = 'Succeeded' }
}

Export-ModuleMember -Function Get-WinToolsMenuItem, Get-WinToolsAppMenuItem, Invoke-WinToolsMenuItem
