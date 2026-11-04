$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$toolRoot = Join-Path $repoRoot 'tools\temporary-save'
$copyScript = Join-Path $toolRoot 'Copy-TemporaryItem.ps1'
$clearScript = Join-Path $toolRoot 'Clear-TemporaryFolder.ps1'
$watchScript = Join-Path $toolRoot 'Start-TemporaryFolderWatcher.ps1'
$registerScript = Join-Path $repoRoot 'setup\Register-TemporarySaveMenu.ps1'
$unregisterScript = Join-Path $repoRoot 'setup\Unregister-TemporarySaveMenu.ps1'
Import-Module (Join-Path $toolRoot 'lib\TemporarySave.psm1') -Force

Describe 'Temporary save with an independent data directory' {
  It 'copies literal Japanese paths containing brackets into the requested folder' {
    $source = Join-Path $TestDrive '案件 [1].txt'
    [IO.File]::WriteAllText($source, 'sample')
    $folder = Join-Path $TestDrive 'saved files'
    $result = & $copyScript -LiteralPath $source -FolderPath $folder
    $result.DirectoryName | Should Be $folder
    $result.Name | Should Match '^【一時保存】＜\d{8}＞案件 \[1\]\.txt$'
    [IO.File]::ReadAllText($result.FullName) | Should Be 'sample'
    Test-Path -LiteralPath $source | Should Be $true
  }
  It 'copies empty and nested directories' {
    $source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'source')
    New-Item -ItemType Directory -Path (Join-Path $source.FullName 'empty') | Out-Null
    [IO.File]::WriteAllText((Join-Path $source.FullName 'file.txt'), 'content')
    $result = & $copyScript -LiteralPath $source.FullName -FolderPath (Join-Path $TestDrive 'copies')
    Test-Path -LiteralPath (Join-Path $result.FullName 'empty') -PathType Container | Should Be $true
    [IO.File]::ReadAllText((Join-Path $result.FullName 'file.txt')) | Should Be 'content'
  }
  It 'does not overwrite a copy made on the same date' {
    $source = Join-Path $TestDrive 'collision.txt'
    [IO.File]::WriteAllText($source, 'original')
    $folder = Join-Path $TestDrive 'collision-target'
    $first = & $copyScript -LiteralPath $source -FolderPath $folder
    [IO.File]::WriteAllText($source, 'updated')
    { & $copyScript -LiteralPath $source -FolderPath $folder } | Should Throw
    [IO.File]::ReadAllText($first.FullName) | Should Be 'original'
  }
  It 'rejects a recursive copy into the source directory' {
    $source = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'recursive')
    { & $copyScript -LiteralPath $source.FullName -FolderPath (Join-Path $source.FullName 'saved') } | Should Throw
    Test-Path -LiteralPath (Join-Path $source.FullName 'saved') | Should Be $false
  }
  It 'does not create the data directory or open Explorer during WhatIf' {
    Mock Start-Process { throw 'Explorer should not start.' }
    $source = Join-Path $TestDrive 'dry-run.txt'
    [IO.File]::WriteAllText($source, 'sample')
    $folder = Join-Path $TestDrive 'dry-run-target'
    & $copyScript -LiteralPath $source -FolderPath $folder -Open -WhatIf
    Test-Path -LiteralPath $folder | Should Be $false
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
  }
  It 'rejects a drive root as a data directory' {
    { Get-TemporaryFolderPath -FolderPath ([IO.Path]::GetPathRoot($TestDrive)) } | Should Throw
  }
  It 'clears only marked direct children and leaves ordinary items untouched' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'clear')
    $marked = Join-Path $folder.FullName '【一時保存】＜20261104＞copy [1].txt'
    [IO.File]::WriteAllText($marked, 'remove')
    $ordinary = New-Item -ItemType Directory -Path (Join-Path $folder.FullName 'ordinary')
    $nested = Join-Path $ordinary.FullName '【一時保存】＜20261104＞keep.txt'
    [IO.File]::WriteAllText($nested, 'keep')
    $malformed = Join-Path $folder.FullName '【一時保存】＜not-a-date＞keep.txt'
    [IO.File]::WriteAllText($malformed, 'keep')
    & $clearScript -FolderPath $folder.FullName -Confirm:$false
    Test-Path -LiteralPath $marked | Should Be $false
    Test-Path -LiteralPath $nested | Should Be $true
    Test-Path -LiteralPath $malformed | Should Be $true
  }
  It 'preserves copies during a cleanup dry run' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'clear-dry-run')
    $marked = Join-Path $folder.FullName '【一時保存】＜20261104＞keep.txt'
    [IO.File]::WriteAllText($marked, 'keep')
    & $clearScript -FolderPath $folder.FullName -WhatIf
    Test-Path -LiteralPath $marked | Should Be $true
  }
  It 'refuses to delete a folder containing a junction and preserves its target' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'clear-links')
    $marked = New-Item -ItemType Directory -Path (Join-Path $folder.FullName '【一時保存】＜20261104＞linked')
    $outside = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'link-target')
    $kept = Join-Path $outside.FullName 'keep.txt'
    [IO.File]::WriteAllText($kept, 'keep')
    $link = Join-Path $marked.FullName 'junction'
    New-Item -ItemType Junction -Path $link -Target $outside.FullName | Out-Null
    try {
      { & $clearScript -FolderPath $folder.FullName -Confirm:$false } | Should Throw
      Test-Path -LiteralPath $kept | Should Be $true
      Test-Path -LiteralPath $marked.FullName | Should Be $true
    } finally {
      # Remove the junction itself without traversing the target before Pester cleanup.
      [IO.Directory]::Delete($link)
    }
  }
  It 'renames files and directories independently of the script location' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'rename')
    $file = Join-Path $folder.FullName '資料 [1].txt'
    [IO.File]::WriteAllText($file, 'sample')
    $directory = New-Item -ItemType Directory -Path (Join-Path $folder.FullName '新規フォルダ')
    Rename-TemporaryItem -LiteralPath $file
    Rename-TemporaryItem -LiteralPath $directory.FullName
    @(Get-ChildItem -LiteralPath $folder.FullName | Where-Object { $_.Name -match '^【一時保存】＜\d{8}＞' }).Count | Should Be 2
  }
  It 'skips temporary, numbered and already prefixed names' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'skipped')
    foreach ($name in @('~lock.txt', '123_keep.txt', '【一時保存】＜20261104＞keep.txt')) {
      $path = Join-Path $folder.FullName $name
      [IO.File]::WriteAllText($path, 'keep')
      Rename-TemporaryItem -LiteralPath $path
      Test-Path -LiteralPath $path | Should Be $true
    }
  }
  It 'detects actual creations, excludes descendants and releases event subscriptions' {
    $folder = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'live-watch')
    $existing = Join-Path $folder.FullName 'existing.txt'
    [IO.File]::WriteAllText($existing, 'existing')
    $child = New-Item -ItemType Directory -Path (Join-Path $folder.FullName 'existing-directory')
    $before = @(Get-EventSubscriber -ErrorAction SilentlyContinue).Count
    $writer = Start-Job -ArgumentList $folder.FullName, $child.FullName -ScriptBlock {
      param($folder, $child)
      Start-Sleep -Seconds 1
      [IO.File]::WriteAllText((Join-Path $folder 'new [1].txt'), 'new')
      [IO.Directory]::CreateDirectory((Join-Path $folder 'new-directory')) | Out-Null
      [IO.File]::WriteAllText((Join-Path $child 'nested.txt'), 'nested')
    }
    try {
      & $watchScript -FolderPath $folder.FullName -NoTray -DurationSeconds 6
      $writer | Wait-Job -Timeout 10 | Out-Null
      $writer.State | Should Be 'Completed'
      $writer | Receive-Job -ErrorAction Stop
      Test-Path -LiteralPath (Join-Path $folder.FullName (Get-TemporaryItemName -Name 'new [1].txt')) | Should Be $true
      Test-Path -LiteralPath (Join-Path $folder.FullName (Get-TemporaryItemName -Name 'new-directory')) -PathType Container | Should Be $true
      Test-Path -LiteralPath $existing | Should Be $true
      Test-Path -LiteralPath (Join-Path $child.FullName 'nested.txt') | Should Be $true
      @(Get-EventSubscriber -ErrorAction SilentlyContinue).Count | Should Be $before
    } finally {
      $writer | Remove-Job -Force
    }
  }
}

Describe 'Temporary save menu registration' {
  It 'registers the relocated tool with a quoted explicit destination for files and folders' {
    Mock New-Item {}
    Mock Set-Item {}
    $folder = Join-Path $TestDrive '保存先 with spaces'
    & $registerScript -FolderPath $folder
    Assert-MockCalled Set-Item -Times 2 -Exactly -Scope It -ParameterFilter {
      $LiteralPath -like '*\command' -and
      $Value.Contains('tools\temporary-save\Copy-TemporaryItem.ps1') -and
      $Value.Contains('-LiteralPath "%1"') -and
      $Value.Contains('-FolderPath "' + $folder + '"')
    }
  }
  It 'does not write registry entries during WhatIf' {
    Mock New-Item {}
    Mock Set-Item {}
    & $registerScript -WhatIf
    Assert-MockCalled New-Item -Times 0 -Exactly -Scope It
    Assert-MockCalled Set-Item -Times 0 -Exactly -Scope It
  }
  It 'removes only the two current-user temporary-save menu keys' {
    Mock Test-Path { $true }
    Mock Remove-Item {}
    & $unregisterScript
    Assert-MockCalled Remove-Item -Times 2 -Exactly -Scope It -ParameterFilter {
      $LiteralPath -eq 'HKCU:\Software\Classes\*\shell\CopyToSelfFolder' -or
      $LiteralPath -eq 'HKCU:\Software\Classes\Directory\shell\CopyToSelfFolder'
    }
  }
}
