$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$shortcutScript = Join-Path $repoRoot 'setup\Register-WinToolsShortcut.ps1'

Describe 'Keyboard shortcut registration' {
  It 'does not create a shortcut during a dry run' {
    $path = Join-Path $TestDrive 'dry-run.lnk'
    & $shortcutScript -ShortcutPath $path -WhatIf
    Test-Path -LiteralPath $path | Should Be $false
  }
  It 'creates a usable shortcut with the hotkey and removes only that file' {
    $path = Join-Path $TestDrive '起動 [1].lnk'
    $neighbor = Join-Path $TestDrive 'keep.txt'
    [IO.File]::WriteAllText($neighbor, 'keep')
    & $shortcutScript -ShortcutPath $path -HotKey 'CTRL+ALT+W' | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $null
    try {
      $shortcut = $shell.CreateShortcut($path)
      $shortcut.TargetPath | Should Match 'powershell.exe$'
      $shortcut.Arguments | Should Be ('-NoProfile -Sta -ExecutionPolicy Bypass -File "' + (Join-Path $repoRoot 'Start-WinTools.ps1') + '"')
      $shortcut.WorkingDirectory | Should Be $repoRoot
      (($shortcut.Hotkey.ToUpperInvariant().Split('+') | Sort-Object) -join '+') | Should Be 'ALT+CTRL+W'
    } finally {
      foreach ($item in @($shortcut, $shell)) {
        if ($item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) }
      }
    }
    & $shortcutScript -ShortcutPath $path -Unregister -WhatIf
    Test-Path -LiteralPath $path | Should Be $true
    & $shortcutScript -ShortcutPath $path -Unregister
    Test-Path -LiteralPath $path | Should Be $false
    Test-Path -LiteralPath $neighbor | Should Be $true
  }
}
