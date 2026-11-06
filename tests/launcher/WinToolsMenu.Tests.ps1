$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$menuModule = Join-Path $repoRoot 'tools\launcher\lib\WinToolsMenu.psm1'
Import-Module $menuModule -Force

Describe 'Tool menu and automatic Excel discovery' {
  BeforeEach {
    $fixture = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $macroFolder = New-Item -ItemType Directory -Path (Join-Path $fixture 'tools\excel-macros') -Force
    $global:WinToolsLauncherFixtureState = @{}
  }
  AfterEach { Remove-Variable -Name WinToolsLauncherFixtureState -Scope Global -ErrorAction SilentlyContinue }
  It 'lists six built-in tools without requesting interactive input' {
    Mock Read-Host { throw 'Listing must not prompt.' }
    $items = @(& (Join-Path $repoRoot 'Start-WinTools.ps1') -List)
    $builtins = @($items | Where-Object Kind -ne 'ExcelMacro')
    $builtins.Count | Should Be 6
    $builtins[0].Kind | Should Be 'FolderCompare'
    $builtins[-1].Kind | Should Be 'TemporaryWatch'
    Assert-MockCalled Read-Host -Times 0 -Exactly -Scope It
  }
  It 'discovers supported workbooks in subfolders and sorts by relative file path' {
    $sub = New-Item -ItemType Directory -Path (Join-Path $macroFolder.FullName '集計')
    foreach ($name in @('z.xlsm', 'a.XLSB', 'tool.xlam', 'template.xltm', 'plain.xlsx', 'source.bas')) {
      [IO.File]::WriteAllText((Join-Path $macroFolder.FullName $name), 'fixture only')
    }
    [IO.File]::WriteAllText((Join-Path $sub.FullName '案件 [1].xlsm'), 'fixture only')
    $items = @(Get-WinToolsMenuItem -RepositoryRoot $fixture)
    $items.Count | Should Be 5
    $items[0].Name | Should Be 'Excel: a.XLSB'
    $items[-1].Name | Should Be 'Excel: 集計\案件 [1].xlsm'
    ($items.Number -join ',') | Should Be '1,2,3,4,5'
  }
  It 'excludes workbook locks, private settings and data folders' {
    foreach ($name in @('~$tool.xlsm', 'tool.local.xlsm', 'tool.secrets.xlsb', 'tool.credentials.xlsm')) {
      [IO.File]::WriteAllText((Join-Path $macroFolder.FullName $name), 'private fixture')
    }
    foreach ($name in @('data', 'input', 'output', 'local', '.hidden')) {
      $folder = New-Item -ItemType Directory -Path (Join-Path $macroFolder.FullName $name)
      [IO.File]::WriteAllText((Join-Path $folder.FullName 'private.xlsm'), 'private fixture')
    }
    $hidden = Join-Path $macroFolder.FullName 'hidden.xlsm'
    [IO.File]::WriteAllText($hidden, 'private fixture')
    [IO.File]::SetAttributes($hidden, [IO.FileAttributes]::Hidden)
    @(Get-WinToolsMenuItem -RepositoryRoot $fixture).Count | Should Be 0
  }
  It 'refreshes additions and removals without changing launcher code' {
    $path = Join-Path $macroFolder.FullName 'added.xlsm'
    @(Get-WinToolsMenuItem -RepositoryRoot $fixture).Count | Should Be 0
    [IO.File]::WriteAllText($path, 'fixture only')
    @(Get-WinToolsMenuItem -RepositoryRoot $fixture).Count | Should Be 1
    [IO.File]::Delete($path)
    @(Get-WinToolsMenuItem -RepositoryRoot $fixture).Count | Should Be 0
  }
  It 'does not scan linked directories' {
    $outside = New-Item -ItemType Directory -Path (Join-Path $fixture 'outside')
    [IO.File]::WriteAllText((Join-Path $outside.FullName 'not-a-tool.xlsm'), 'fixture only')
    $link = Join-Path $macroFolder.FullName 'linked'
    New-Item -ItemType Junction -Path $link -Target $outside.FullName | Out-Null
    try { @(Get-WinToolsMenuItem -RepositoryRoot $fixture).Count | Should Be 0 }
    finally { [IO.Directory]::Delete($link) }
  }
  It 'opens only the selected workbook using its literal full path' {
    Mock Start-Process {} -ModuleName WinToolsMenu
    $path = Join-Path $macroFolder.FullName '集計 [1] & tool.xlsm'
    [IO.File]::WriteAllText($path, 'fixture only')
    $item = @(Get-WinToolsMenuItem -RepositoryRoot $fixture)[0]
    Invoke-WinToolsMenuItem -Item $item
    Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ModuleName WinToolsMenu -ParameterFilter { $FilePath -eq $path -and -not $ArgumentList }
  }
  It 'does not launch a workbook deleted after listing' {
    Mock Start-Process {} -ModuleName WinToolsMenu
    $path = Join-Path $macroFolder.FullName 'removed.xlsm'
    [IO.File]::WriteAllText($path, 'fixture only')
    $item = @(Get-WinToolsMenuItem -RepositoryRoot $fixture)[0]
    [IO.File]::Delete($path)
    { Invoke-WinToolsMenuItem -Item $item } | Should Throw
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It -ModuleName WinToolsMenu
  }
  It 'launches a macro template with the New verb rather than editing its source' {
    Mock Start-Process {} -ModuleName WinToolsMenu
    $path = Join-Path $macroFolder.FullName 'template.xltm'
    [IO.File]::WriteAllText($path, 'fixture only')
    $item = @(Get-WinToolsMenuItem -RepositoryRoot $fixture)[0]
    Invoke-WinToolsMenuItem -Item $item
    Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ModuleName WinToolsMenu -ParameterFilter {
      $FilePath -eq $path -and $Verb -eq 'New'
    }
  }
  It 'passes quoted paths as literal parameters and enables the requested hash comparison' {
    $leftInput = (New-Item -ItemType Directory -Path (Join-Path $fixture '左 [1] & folder')).FullName
    $rightInput = (New-Item -ItemType Directory -Path (Join-Path $fixture 'right folder')).FullName
    $global:WinToolsLauncherFixtureState.Left = $leftInput
    $global:WinToolsLauncherFixtureState.Right = $rightInput
    $fakeTool = Join-Path $fixture 'compare.ps1'
    $code = @'
param($LeftPath, $RightPath, $CompareBy, [switch]$DifferencesOnly)
[IO.File]::WriteAllText((Join-Path $RightPath 'invocation.txt'), ($LeftPath + "`n" + $CompareBy + "`n" + $DifferencesOnly))
'@
    [IO.File]::WriteAllText($fakeTool, $code)
    Mock Read-Host {
      if ($Prompt -like '左側*') { '"' + $global:WinToolsLauncherFixtureState.Left + '"' }
      elseif ($Prompt -like '右側*') { "'" + $global:WinToolsLauncherFixtureState.Right + "'" }
      else { 'y' }
    } -ModuleName WinToolsMenu
    Invoke-WinToolsMenuItem -Item ([pscustomobject]@{ Kind = 'FolderCompare'; Path = $fakeTool })
    $arguments = [IO.File]::ReadAllLines((Join-Path $rightInput 'invocation.txt'))
    $arguments[0] | Should Be $leftInput
    $arguments[1] | Should Be 'Hash'
    $arguments[2] | Should Be 'True'
  }
  It 'starts the tray watcher in a hidden process with an explicit quoted data folder' {
    $watchInput = Join-Path $fixture '保存先 with spaces'
    $global:WinToolsLauncherFixtureState.Watch = $watchInput
    Mock Read-Host { $global:WinToolsLauncherFixtureState.Watch } -ModuleName WinToolsMenu
    Mock Start-Process {} -ModuleName WinToolsMenu
    $item = Get-WinToolsMenuItem -RepositoryRoot $repoRoot | Where-Object Kind -eq 'TemporaryWatch'
    Invoke-WinToolsMenuItem -Item $item
    Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ModuleName WinToolsMenu -ParameterFilter {
      $FilePath -eq 'powershell.exe' -and $WindowStyle -eq 'Hidden' -and
      $ArgumentList -contains ('"' + $watchInput + '"')
    }
  }
  It 'keeps built-in entries available if a macro directory cannot be read' {
    $fileFolder = New-Item -ItemType Directory -Path (Join-Path $fixture 'tools\files') -Force
    [IO.File]::WriteAllText((Join-Path $fileFolder.FullName 'Compare-Folders.ps1'), '# fixture')
    Mock Get-ChildItem { throw 'Access denied' } -ModuleName WinToolsMenu
    $warnings = @()
    $items = @(Get-WinToolsMenuItem -RepositoryRoot $fixture -WarningVariable warnings -WarningAction SilentlyContinue)
    $items.Count | Should Be 1
    $items[0].Kind | Should Be 'FolderCompare'
    $warnings.Count | Should Be 1
  }
  It 'exits when zero is entered' {
    Mock Read-Host { '0' }
    & (Join-Path $repoRoot 'Start-WinTools.ps1')
    Assert-MockCalled Read-Host -Times 1 -Exactly -Scope It
  }
  It 'returns to the prompt after an invalid number' {
    $script:promptCount = 0
    Mock Read-Host { $script:promptCount++; if ($script:promptCount -eq 1) { '999' } else { '0' } }
    $warnings = @()
    & (Join-Path $repoRoot 'Start-WinTools.ps1') -WarningVariable warnings -WarningAction SilentlyContinue
    $warnings.Count | Should Be 1
    Assert-MockCalled Read-Host -Times 2 -Exactly -Scope It
  }
}

Describe 'Git policy for future Excel tools' {
  It 'allows macro tools to be versioned only in the designated directory' {
    foreach ($extension in @('xlsm', 'xlsb', 'xlam', 'xltm')) {
      git -C $repoRoot check-ignore -q --no-index -- "tools/excel-macros/group/tool.$extension" 2>$null
      $LASTEXITCODE | Should Be 1
      git -C $repoRoot check-ignore -q --no-index -- "work.$extension" 2>$null
      $LASTEXITCODE | Should Be 0
    }
  }
  It 'continues to exclude workbook locks, ordinary workbooks and private output' {
    foreach ($path in @('~$tool.xlsm', 'plain.xlsx', '.env.xlsm', '.hidden/tool.xlsm', 'data/private.xlsm', 'group/input/private.xlsb', 'output/report.xlsm', 'local/tool.xlsm', 'tool.local.xlsm', 'tool.secrets.xlsm', 'tool.credentials.xlsm')) {
      git -C $repoRoot check-ignore -q --no-index -- ('tools/excel-macros/' + $path) 2>$null
      $LASTEXITCODE | Should Be 0
    }
  }
}
