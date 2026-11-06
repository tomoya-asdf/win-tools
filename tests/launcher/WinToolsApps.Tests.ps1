$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'tools\launcher\lib\WinToolsMenu.psm1') -Force

Describe 'User application discovery and launching' {
  BeforeEach {
    $fixture = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $apps = (New-Item -ItemType Directory -Path (Join-Path $fixture 'apps') -Force).FullName
    $global:WinToolsAppsFixture = @{}
    Mock Clear-Host {} -ModuleName WinToolsMenu
    Mock Write-WinToolsHistory { $global:WinToolsAppsFixture.Record = $Record } -ModuleName WinToolsMenu
    Mock Start-Process {} -ModuleName WinToolsMenu
  }
  AfterEach { Remove-Variable WinToolsAppsFixture -Scope Global -ErrorAction SilentlyContinue }

  It 'lists directories and associated files one level at a time even without Excel tools' {
    $group = New-Item -ItemType Directory -Path (Join-Path $apps '仕事')
    [IO.File]::WriteAllText((Join-Path $group.FullName 'nested.exe'), 'fixture')
    foreach ($name in @('tool.exe','browser.lnk','website.url','manual.txt','.hidden.exe')) {
      [IO.File]::WriteAllText((Join-Path $apps $name), 'fixture')
    }
    $items = @(Get-WinToolsMenuItem -RepositoryRoot $fixture)
    $items.Count | Should Be 5
    $items[0].Kind | Should Be 'AppFolder'
    $items[0].Name | Should Be 'フォルダ: 仕事'
    ($items.Number -join ',') | Should Be '1,2,3,4,5'
    @($items | Where-Object Name -like '*nested.exe').Count | Should Be 0
    $children = @(Get-WinToolsAppMenuItem -Directory $group.FullName -Root $apps)
    $children.Count | Should Be 1
    $children[0].Name | Should Be 'アプリ: nested.exe'
  }
  It 'refreshes added and removed applications and skips hidden items' {
    $path = Join-Path $apps 'new.exe'
    @(Get-WinToolsAppMenuItem -Directory $apps).Count | Should Be 0
    [IO.File]::WriteAllText($path, 'fixture')
    @(Get-WinToolsAppMenuItem -Directory $apps).Count | Should Be 1
    [IO.File]::SetAttributes($path, [IO.FileAttributes]::Hidden)
    @(Get-WinToolsAppMenuItem -Directory $apps).Count | Should Be 0
    [IO.File]::SetAttributes($path, [IO.FileAttributes]::Normal)
    [IO.File]::Delete($path)
    @(Get-WinToolsAppMenuItem -Directory $apps).Count | Should Be 0
  }
  It 'opens files using their literal full paths and records only the launch result' {
    foreach ($name in @('仕事 [1] & app.exe','shortcut.lnk','website.url','readme.txt')) {
      [IO.File]::WriteAllText((Join-Path $apps $name), 'fixture')
    }
    foreach ($item in @(Get-WinToolsAppMenuItem -Directory $apps)) {
      Invoke-WinToolsMenuItem -Item $item
      $global:WinToolsAppsFixture.Record.Status | Should Be 'Launched'
      $global:WinToolsAppsFixture.Record.ToolPath | Should Be $item.Path
    }
    Assert-MockCalled Start-Process -Times 4 -Exactly -Scope It -ModuleName WinToolsMenu -ParameterFilter {
      $FilePath.StartsWith($apps + '\') -and $WorkingDirectory -eq $apps -and -not $ArgumentList
    }
  }
  It 'records failure if a file is deleted or has no usable association' {
    $path = Join-Path $apps 'removed.exe'
    [IO.File]::WriteAllText($path, 'fixture')
    $item = @(Get-WinToolsAppMenuItem -Directory $apps)[0]
    [IO.File]::Delete($path)
    { Invoke-WinToolsMenuItem -Item $item } | Should Throw
    $global:WinToolsAppsFixture.Record.Status | Should Be 'Failed'
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It -ModuleName WinToolsMenu
    [IO.File]::WriteAllText($path, 'fixture')
    Mock Start-Process { throw 'No application is associated.' } -ModuleName WinToolsMenu
    { Invoke-WinToolsMenuItem -Item $item } | Should Throw 'No application is associated.'
    $global:WinToolsAppsFixture.Record.Error | Should Be 'No application is associated.'
  }
  It 'discovers and launches files several levels below the registration root' {
    $deep = (New-Item -ItemType Directory -Path (Join-Path $apps 'one\two\three') -Force).FullName
    $path = Join-Path $deep 'nested.lnk'
    [IO.File]::WriteAllText($path, 'fixture')
    $item = @(Get-WinToolsAppMenuItem -Directory $deep -Root $apps)[0]
    Invoke-WinToolsMenuItem -Item $item
    Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ModuleName WinToolsMenu -ParameterFilter {
      $FilePath -eq $path -and $WorkingDirectory -eq $deep
    }
    $global:WinToolsAppsFixture.Record.Status | Should Be 'Launched'
  }
  It 'does not follow junctions or a directory replaced with a junction after selection' {
    $outside = New-Item -ItemType Directory -Path (Join-Path $fixture 'outside')
    [IO.File]::WriteAllText((Join-Path $outside.FullName 'outside.exe'), 'fixture')
    $group = New-Item -ItemType Directory -Path (Join-Path $apps 'group')
    $link = $group.FullName
    [IO.Directory]::Delete($link)
    New-Item -ItemType Junction -Path $link -Target $outside.FullName | Out-Null
    try {
      @(Get-WinToolsAppMenuItem -Directory $apps).Count | Should Be 0
      { Get-WinToolsAppMenuItem -Directory $link -Root $apps } | Should Throw
      $item = [pscustomobject]@{ Name = 'app'; Kind = 'UserApp'; Path = (Join-Path $link 'outside.exe'); AppRoot = $apps }
      { Invoke-WinToolsMenuItem -Item $item } | Should Throw
      Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It -ModuleName WinToolsMenu
    } finally { [IO.Directory]::Delete($link) }
  }
  It 'rejects paths outside the registration folder including sibling prefix matches' {
    $sibling = New-Item -ItemType Directory -Path (Join-Path $fixture 'apps-other')
    { Get-WinToolsAppMenuItem -Directory $sibling.FullName -Root $apps } | Should Throw
    { Get-WinToolsAppMenuItem -Directory (Join-Path $apps '..') -Root $apps } | Should Throw
  }
  It 'keeps built-in tools available when app registration is unreadable' {
    $toolFolder = New-Item -ItemType Directory -Path (Join-Path $fixture 'tools\files') -Force
    [IO.File]::WriteAllText((Join-Path $toolFolder.FullName 'Compare-Folders.ps1'), '# fixture')
    Mock Get-ChildItem { throw 'Access denied' } -ModuleName WinToolsMenu -ParameterFilter { $LiteralPath -eq $apps }
    $warnings = @()
    $items = @(Get-WinToolsMenuItem -RepositoryRoot $fixture -WarningVariable warnings -WarningAction SilentlyContinue)
    $items.Count | Should Be 1
    $items[0].Kind | Should Be 'FolderCompare'
    $warnings.Count | Should Be 1
  }
}

Describe 'Application submenu navigation' {
  BeforeEach {
    $appRoot = Join-Path $repoRoot 'apps'
    $group = Join-Path $appRoot 'group'
    $nested = Join-Path $group 'nested'
    $global:WinToolsAppsNavigation = @{
      Root = $appRoot; Group = $group; Nested = $nested
      Choices = [Collections.Generic.Queue[string]]::new()
    }
    Mock Clear-Host {}
    Mock Get-WinToolsMenuItem {
      [pscustomobject]@{ Number = 1; Name = 'group'; Kind = 'AppFolder'; Path = $global:WinToolsAppsNavigation.Group }
    }
    Mock Get-WinToolsAppMenuItem {
      if ($Directory -eq $global:WinToolsAppsNavigation.Group) {
        [pscustomobject]@{ Number = 1; Name = 'nested'; Kind = 'AppFolder'; Path = $global:WinToolsAppsNavigation.Nested }
      } else {
        [pscustomobject]@{ Number = 1; Name = 'app'; Kind = 'UserApp'; Path = (Join-Path $Directory 'app.exe') }
      }
    }
    Mock Read-Host {
      if (-not $global:WinToolsAppsNavigation.Choices.Count) { throw 'Unexpected prompt.' }
      $global:WinToolsAppsNavigation.Choices.Dequeue()
    }
    Mock Invoke-WinToolsMenuItem { $global:WinToolsAppsNavigation.Launched = $Item }
  }
  AfterEach { Remove-Variable WinToolsAppsNavigation -Scope Global -ErrorAction SilentlyContinue }

  It 'navigates nested folders, returns to the same submenu after launch, and goes back one level at a time' {
    foreach ($choice in @('1','1','1','','0','0','0')) { $global:WinToolsAppsNavigation.Choices.Enqueue($choice) }
    & (Join-Path $repoRoot 'Start-WinTools.ps1')
    $global:WinToolsAppsNavigation.Launched.Kind | Should Be 'UserApp'
    $global:WinToolsAppsNavigation.Launched.Path | Should Be (Join-Path $nested 'app.exe')
    Assert-MockCalled Invoke-WinToolsMenuItem -Times 1 -Exactly -Scope It
    Assert-MockCalled Get-WinToolsAppMenuItem -Times 2 -Exactly -Scope It -ParameterFilter { $Directory -eq $nested }
    Assert-MockCalled Get-WinToolsAppMenuItem -Times 2 -Exactly -Scope It -ParameterFilter { $Directory -eq $group }
  }
  It 'shows an empty submenu and allows returning without launching anything' {
    Mock Get-WinToolsAppMenuItem {}
    foreach ($choice in @('1','0','0')) { $global:WinToolsAppsNavigation.Choices.Enqueue($choice) }
    & (Join-Path $repoRoot 'Start-WinTools.ps1')
    Assert-MockCalled Invoke-WinToolsMenuItem -Times 0 -Exactly -Scope It
  }
  It 'returns to the main menu if the current folder disappears' {
    Mock Get-WinToolsAppMenuItem { throw 'Folder no longer exists.' }
    foreach ($choice in @('1','0')) { $global:WinToolsAppsNavigation.Choices.Enqueue($choice) }
    $warnings = @()
    & (Join-Path $repoRoot 'Start-WinTools.ps1') -WarningVariable warnings -WarningAction SilentlyContinue
    $warnings[0].ToString() | Should Match 'Folder no longer exists'
    Assert-MockCalled Invoke-WinToolsMenuItem -Times 0 -Exactly -Scope It
  }
}

Describe 'Private application registration in Git' {
  It 'ignores the entire apps tree regardless of filename or extension' {
    foreach ($path in @('apps/tool.exe','apps/tool.lnk','apps/work/link.url','apps/notes.txt','apps/nested/.gitkeep','apps/tool.xlsm')) {
      git -C $repoRoot check-ignore -q --no-index -- $path 2>$null
      $LASTEXITCODE | Should Be 0
    }
    @(git -C $repoRoot ls-files -- apps).Count | Should Be 0
  }
}
