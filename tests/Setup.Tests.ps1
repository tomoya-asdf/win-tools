$repoRoot = Split-Path -Parent $PSScriptRoot
$registerScript = Join-Path $repoRoot 'setup\Register-WinToolsMenus.ps1'

Describe 'Optional menu registration without home deployment' {
  It 'does not change the environment when no menu is selected' {
    Mock New-Item {}
    Mock Set-Item {}
    Mock Start-Process {}
    & $registerScript
    Assert-MockCalled New-Item -Times 0 -Exactly -Scope It
    Assert-MockCalled Set-Item -Times 0 -Exactly -Scope It
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
  }
  It 'registers the renamed copy tool with an explicit data folder' {
    Mock New-Item {}
    Mock Set-Item {}
    Mock Start-Process { throw 'Elevation is unnecessary.' }
    & $registerScript -RegisterContextMenu -TemporaryFolderPath (Join-Path $TestDrive 'saved')
    Assert-MockCalled Set-Item -Times 2 -Exactly -Scope It -ParameterFilter {
      $LiteralPath -like '*\command' -and $Value.Contains('Copy-TemporaryItem.ps1')
    }
    Assert-MockCalled New-Item -Times 2 -Exactly -Scope It -ParameterFilter { $Path -like 'HKCU:*' }
    Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
  }
  It 'still registers Excel read-only entries without creating data directories' {
    Mock New-Item {}
    Mock Set-Item {}
    Mock New-ItemProperty {}
    Mock Get-ItemProperty { [pscustomobject]@{ '(default)' = 'C:\Test\EXCEL.EXE' } }
    Mock Test-Path { $true } -ParameterFilter { $LiteralPath -eq 'C:\Test\EXCEL.EXE' }
    & $registerScript -RegisterExcelReadOnly
    Assert-MockCalled Set-Item -Times 5 -Exactly -Scope It -ParameterFilter {
      $LiteralPath -like '*\command' -and $Value.Contains('C:\Test\EXCEL.EXE') -and $Value.Contains('/r "%1"')
    }
    Assert-MockCalled New-Item -Times 10 -Exactly -Scope It -ParameterFilter { $Path -like 'HKCU:*' }
    Assert-MockCalled New-Item -Times 0 -Exactly -Scope It -ParameterFilter { $ItemType -eq 'Directory' }
  }
  It 'does not write either menu during a dry run' {
    Mock New-Item {}
    Mock Set-Item {}
    Mock New-ItemProperty {}
    Mock Get-ItemProperty { [pscustomobject]@{ '(default)' = 'C:\Test\EXCEL.EXE' } }
    Mock Test-Path { $true } -ParameterFilter { $LiteralPath -eq 'C:\Test\EXCEL.EXE' }
    & $registerScript -RegisterContextMenu -RegisterExcelReadOnly -WhatIf
    Assert-MockCalled New-Item -Times 0 -Exactly -Scope It
    Assert-MockCalled Set-Item -Times 0 -Exactly -Scope It
    Assert-MockCalled New-ItemProperty -Times 0 -Exactly -Scope It
  }
  It 'does not create a shortcut during a dry run through the combined setup' {
    Mock New-Object { throw 'A dry run must not instantiate WScript.Shell.' } -ParameterFilter { $ComObject -eq 'WScript.Shell' }
    & $registerScript -RegisterShortcut -WhatIf
    Assert-MockCalled New-Object -Times 0 -Exactly -Scope It -ParameterFilter { $ComObject -eq 'WScript.Shell' }
  }
}
