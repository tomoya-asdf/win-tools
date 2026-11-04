$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$compareScript = Join-Path $repoRoot 'tools\files\Compare-Folders.ps1'

Describe 'Folder comparison' {
  BeforeEach {
    $caseRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $left = New-Item -ItemType Directory -Path (Join-Path $caseRoot 'left') -Force
    $right = New-Item -ItemType Directory -Path (Join-Path $caseRoot 'right') -Force
  }
  It 'reports files present on only one side using relative paths' {
    [IO.File]::WriteAllText((Join-Path $left.FullName '左 [1].txt'), 'left')
    [IO.File]::WriteAllText((Join-Path $right.FullName 'right.txt'), 'right')
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName)
    $rows.Count | Should Be 2
    ($rows | Where-Object RelativePath -eq '左 [1].txt').Status | Should Be 'OnlyLeft'
    ($rows | Where-Object RelativePath -eq 'right.txt').Status | Should Be 'OnlyRight'
  }
  It 'compares same-sized files without claiming their content matches' {
    [IO.File]::WriteAllText((Join-Path $left.FullName 'same.txt'), 'AAAA')
    [IO.File]::WriteAllText((Join-Path $right.FullName 'same.txt'), 'BBBB')
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName)
    $rows[0].Status | Should Be 'SameSize'
    $rows[0].LeftHash | Should BeNullOrEmpty
    @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash)[0].Status | Should Be 'DifferentContent'
  }
  It 'returns SHA-256 values for matching content and detects size differences' {
    foreach ($root in @($left.FullName, $right.FullName)) {
      [IO.File]::WriteAllText((Join-Path $root 'identical.txt'), 'same content')
    }
    [IO.File]::WriteAllText((Join-Path $left.FullName 'size.txt'), 'short')
    [IO.File]::WriteAllText((Join-Path $right.FullName 'size.txt'), 'longer content')
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash)
    $same = $rows | Where-Object RelativePath -eq 'identical.txt'
    $same.Status | Should Be 'SameContent'
    $same.LeftHash.Length | Should Be 64
    $same.LeftHash | Should Be $same.RightHash
    ($rows | Where-Object RelativePath -eq 'size.txt').Status | Should Be 'DifferentSize'
  }
  It 'includes empty directories and nested files' {
    foreach ($root in @($left.FullName, $right.FullName)) {
      New-Item -ItemType Directory -Path (Join-Path $root 'empty') | Out-Null
      New-Item -ItemType Directory -Path (Join-Path $root 'nested') | Out-Null
      [IO.File]::WriteAllText((Join-Path $root 'nested\file.txt'), 'same')
    }
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash)
    ($rows | Where-Object RelativePath -eq 'empty').Status | Should Be 'Directory'
    ($rows | Where-Object RelativePath -eq 'nested\file.txt').Status | Should Be 'SameContent'
    @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash -DifferencesOnly).Count | Should Be 0
  }
  It 'reports a file versus directory mismatch' {
    [IO.File]::WriteAllText((Join-Path $left.FullName 'entry'), 'file')
    New-Item -ItemType Directory -Path (Join-Path $right.FullName 'entry') | Out-Null
    @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName)[0].Status | Should Be 'TypeMismatch'
  }
  It 'treats case variations of a Windows relative path as the same entry' {
    [IO.File]::WriteAllText((Join-Path $left.FullName 'REPORT.txt'), 'same')
    [IO.File]::WriteAllText((Join-Path $right.FullName 'report.TXT'), 'same')
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash)
    $rows.Count | Should Be 1
    $rows[0].Status | Should Be 'SameContent'
  }
  It 'does not follow junctions or falsely label their unseen contents as missing' {
    $target = New-Item -ItemType Directory -Path (Join-Path $caseRoot 'outside')
    [IO.File]::WriteAllText((Join-Path $target.FullName 'outside.txt'), 'outside')
    $link = Join-Path $left.FullName 'linked'
    New-Item -ItemType Junction -Path $link -Target $target.FullName | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $right.FullName 'linked') | Out-Null
    [IO.File]::WriteAllText((Join-Path $right.FullName 'linked\child.txt'), 'child')
    try {
      $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -DifferencesOnly)
      ($rows | Where-Object RelativePath -eq 'linked').Status | Should Be 'SkippedLink'
      ($rows | Where-Object RelativePath -eq 'linked\child.txt').Status | Should Be 'NotCompared'
      @($rows | Where-Object RelativePath -like '*outside.txt').Count | Should Be 0
    } finally { [IO.Directory]::Delete($link) }
  }
  It 'keeps unreadable folders and their unmatched descendants inconclusive' {
    $denied = New-Item -ItemType Directory -Path (Join-Path $caseRoot 'denied-root')
    Mock Get-ChildItem { throw 'Access denied' } -ParameterFilter { $LiteralPath -like '*\denied-root' }
    [IO.File]::WriteAllText((Join-Path $right.FullName 'uncertain.txt'), 'sample')
    $rows = @(& $compareScript -LeftPath $denied.FullName -RightPath $right.FullName)
    ($rows | Where-Object RelativePath -eq '.').Status | Should Be 'ReadError'
    ($rows | Where-Object RelativePath -eq 'uncertain.txt').Status | Should Be 'NotCompared'
  }
  It 'keeps hash read errors visible even when only differences are requested' {
    foreach ($root in @($left.FullName, $right.FullName)) { [IO.File]::WriteAllText((Join-Path $root 'locked.txt'), 'same') }
    Mock Get-FileHash { throw 'Cannot read file' } -ParameterFilter { $LiteralPath -like '*\left\locked.txt' }
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash -DifferencesOnly)
    $rows.Count | Should Be 1
    $rows[0].Status | Should Be 'ReadError'
    $rows[0].Note | Should Match 'Cannot read file'
  }
  It 'does not report content equality when a file changes during hashing' {
    foreach ($root in @($left.FullName, $right.FullName)) { [IO.File]::WriteAllText((Join-Path $root 'changing.txt'), 'same') }
    Mock Get-FileHash {
      [IO.File]::WriteAllText($LiteralPath, 'changed during hashing')
      [pscustomobject]@{ Hash = 'fake-hash' }
    } -ParameterFilter { $LiteralPath -like '*\left\changing.txt' }
    $rows = @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName -CompareBy Hash)
    $rows[0].Status | Should Be 'ReadError'
    $rows[0].Note | Should Match '変更'
  }
  It 'includes hidden files' {
    $path = Join-Path $left.FullName 'hidden.txt'
    [IO.File]::WriteAllText($path, 'hidden')
    [IO.File]::SetAttributes($path, [IO.FileAttributes]::Hidden)
    @(& $compareScript -LeftPath $left.FullName -RightPath $right.FullName)[0].Status | Should Be 'OnlyLeft'
  }
  It 'rejects a file or missing path as a comparison root' {
    $file = Join-Path $left.FullName 'not-a-folder.txt'
    [IO.File]::WriteAllText($file, 'file')
    { & $compareScript -LeftPath $file -RightPath $right.FullName } | Should Throw
    { & $compareScript -LeftPath (Join-Path $left.FullName 'missing') -RightPath $right.FullName } | Should Throw
  }
  It 'does not alter contents or write times' {
    $path = Join-Path $left.FullName 'read-only.txt'
    [IO.File]::WriteAllText($path, 'unchanged')
    $before = (Get-Item -LiteralPath $path).LastWriteTimeUtc
    & $compareScript -LeftPath $left.FullName -RightPath $left.FullName -CompareBy Hash | Out-Null
    [IO.File]::ReadAllText($path) | Should Be 'unchanged'
    (Get-Item -LiteralPath $path).LastWriteTimeUtc | Should Be $before
  }
}
