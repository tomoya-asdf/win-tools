<#
.SYNOPSIS
現在のユーザーの一時保存用右クリックメニューを解除します（管理者不要）。
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param()
foreach ($type in @('*', 'Directory')) {
  $key = "HKCU:\Software\Classes\$type\shell\CopyToSelfFolder"
  if ((Test-Path -LiteralPath $key) -and $PSCmdlet.ShouldProcess($key, 'Unregister temporary save context menu')) {
    Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction Stop
  }
}
