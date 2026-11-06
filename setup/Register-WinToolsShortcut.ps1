<#
.SYNOPSIS
Win Toolsをショートカットキーで起動できるよう、現在のユーザーに登録します。
.DESCRIPTION
既定はスタートメニューのWin Tools.lnkへCtrl+Alt+Wを割り当てます。
常駐プログラムや管理者権限は不要です。リポジトリ移動後は再登録してください。
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [ValidatePattern('^(CTRL\+ALT|CTRL\+SHIFT|ALT\+SHIFT)\+[A-Z0-9]$')]
  [string]$HotKey = 'CTRL+ALT+W',
  [string]$ShortcutPath = (Join-Path ([Environment]::GetFolderPath('Programs')) 'Win Tools.lnk'),
  [switch]$Unregister
)
$shortcutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ShortcutPath)
if ([IO.Path]::GetExtension($shortcutFile) -ne '.lnk') { throw 'ショートカットの拡張子は.lnkにしてください。' }
if ($Unregister) {
  if ((Test-Path -LiteralPath $shortcutFile) -and $PSCmdlet.ShouldProcess($shortcutFile, 'Unregister Win Tools shortcut')) {
    Remove-Item -LiteralPath $shortcutFile -Force -ErrorAction Stop
  }
  return
}
$repoRoot = Split-Path -Parent $PSScriptRoot
$launcher = Join-Path $repoRoot 'Start-WinTools.ps1'
if (-not (Test-Path -LiteralPath $launcher -PathType Leaf)) { throw 'Start-WinTools.ps1が見つかりません。' }
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not $PSCmdlet.ShouldProcess($shortcutFile, ('Register Win Tools shortcut (' + $HotKey + ')'))) { return }
$shell = $null
$shortcut = $null
try {
  [void][IO.Directory]::CreateDirectory((Split-Path -Parent $shortcutFile))
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($shortcutFile)
  $shortcut.TargetPath = $powershell
  $shortcut.Arguments = '-NoProfile -Sta -ExecutionPolicy Bypass -File "' + $launcher + '"'
  $shortcut.WorkingDirectory = $repoRoot
  $shortcut.Description = 'Win Tools 起動メニュー'
  $shortcut.Hotkey = $HotKey.ToUpperInvariant()
  $shortcut.WindowStyle = 1
  $shortcut.Save()
} finally {
  foreach ($item in @($shortcut, $shell)) {
    if ($item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
      [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)
    }
  }
}
Get-Item -LiteralPath $shortcutFile
