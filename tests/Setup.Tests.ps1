$repoRoot = Split-Path -Parent $PSScriptRoot
$linkScript = Join-Path $repoRoot 'setup\New-HomeLinks.ps1'
$installScript = Join-Path $repoRoot 'setup\Install-WinTools.ps1'

Describe 'Setup entry points after directory reorganization' {
  It 'resolves the bundled home directory in a dry run without creating links' {
    Mock New-Item { throw 'A dry run must not create anything.' }
    & $linkScript -TargetRoot $TestDrive -WhatIf
    @(Get-ChildItem -LiteralPath $TestDrive -Force).Count | Should Be 0
    Assert-MockCalled New-Item -Times 0 -Exactly -Scope It
  }
  It 'finds the sibling link script without changing the user environment' {
    Mock Start-Process { throw 'No registration was requested.' }
    & $installScript -MakeLinks:$false
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
  }
  It 'delegates menu registration to setup without elevation' {
    Mock New-Item {}
    Mock Set-Item {}
    Mock Start-Process { throw 'Elevation is unnecessary.' }
    & $installScript -MakeLinks:$false -RegisterContextMenu -TemporaryFolderPath (Join-Path $TestDrive 'saved')
    Assert-MockCalled Set-Item -Times 2 -Exactly -Scope It -ParameterFilter {
      $LiteralPath -like '*\command' -and $Value.Contains('Copy-ToTemporaryFolder.ps1')
    }
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
  }
}
