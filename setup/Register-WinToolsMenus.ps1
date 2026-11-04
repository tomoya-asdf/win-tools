<#
.SYNOPSIS
現在のユーザーへ一時保存・Excel読み取り専用の右クリックメニューを登録します。
.DESCRIPTION
登録するメニューをオプションで指定します。引数なしの場合は使い方を表示します。
.PARAMETER RegisterContextMenu
一時保存のメニューを登録します。
.PARAMETER RegisterExcelReadOnly
Excelを読み取り専用で開くメニューを登録します。
.PARAMETER TemporaryFolderPath
一時保存先。既定はユーザープロファイル直下の一時保存フォルダです。
.EXAMPLE
.\setup\Register-WinToolsMenus.ps1 -RegisterContextMenu -RegisterExcelReadOnly -WhatIf
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [switch]$RegisterContextMenu,
  [switch]$RegisterExcelReadOnly,
  [string]$TemporaryFolderPath = (Join-Path $env:USERPROFILE '一時保存フォルダ')
)

if (-not $RegisterContextMenu -and -not $RegisterExcelReadOnly) {
  Write-Host '登録するメニューを -RegisterContextMenu / -RegisterExcelReadOnly で指定してください。'
  return
}
$scriptDir = Split-Path -Parent $PSCommandPath
if ($RegisterContextMenu) {
  & (Join-Path $scriptDir 'Register-TemporarySaveMenu.ps1') -FolderPath $TemporaryFolderPath
}

if ($RegisterExcelReadOnly) {
  # Adds "Excel(読み取り専用)で開く" to right-click menu for Excel-related extensions under HKCU.
  # No elevation required.

  function Get-ExcelExecutablePath {
    $candidates = @(
      'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe',
      'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe'
    )
    foreach ($key in $candidates) {
      try {
        $p = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop)."(default)"
      } catch { $p = $null }
      if (-not $p) {
        try { $p = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).'(Default)' } catch { }
      }
      if (-not $p) {
        try {
          $pathDir = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).Path
          if ($pathDir) { $p = Join-Path $pathDir 'EXCEL.EXE' }
        } catch { }
      }
      if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    # fallback: try PATH lookup
    try {
      $where = & where.exe excel 2>$null
      if ($where) {
        if ($where -is [array]) { $first = $where[0] } else { $first = $where }
        if ($first -and (Test-Path -LiteralPath $first)) { return $first }
      }
    } catch { }
    return $null
  }

  $excelExe = Get-ExcelExecutablePath
  $iconValue = if ($excelExe) { "$excelExe,0" } else { 'excel.exe,0' }
  $extensions = @('.xlsx', '.xlsm', '.xls', '.xlsb', '.csv')

  foreach ($ext in $extensions) {
    try {
      $baseKey = "HKCU:\\Software\\Classes\\SystemFileAssociations\\$ext\\shell\\OpenReadOnly"
      if (-not $PSCmdlet.ShouldProcess($baseKey, 'Register Excel read-only context menu')) { continue }
      New-Item -Path $baseKey -Force | Out-Null
      New-Item -Path (Join-Path $baseKey 'command') -Force | Out-Null

      # Display name
      Set-Item -LiteralPath $baseKey -Value 'Excel(読み取り専用)で開く'
      # Icon for the menu
      New-ItemProperty -LiteralPath $baseKey -Name 'Icon' -Value $iconValue -PropertyType String -Force | Out-Null

      # Command: prefer explicit Excel path if found; fallback to App Paths/Path lookup.
      $cmdKey = Join-Path $baseKey 'command'
      if ($excelExe) {
        # Use a format string to avoid PowerShell escaping issues
        $cmdVal = ('cmd.exe /c start "" "{0}" /r "%1"' -f $excelExe)
      } else {
        $cmdVal = 'cmd.exe /c start "" excel /r "%1"'
      }
      Set-Item -LiteralPath $cmdKey -Value $cmdVal

      Write-Host ("Registered Excel read-only menu for {0}" -f $ext)
    } catch {
      Write-Warning ("Failed to register Excel read-only for {0}: {1}" -f $ext, $_.Exception.Message)
    }
  }
}

Write-Verbose 'メニュー登録の処理を終了しました。'
