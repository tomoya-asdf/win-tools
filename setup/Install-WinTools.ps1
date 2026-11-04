<#
.
.SYNOPSIS
Installs this repo into a Windows user environment.

.DESCRIPTION
- Creates links from `home` to the user's profile (junctions/symlinks).
- Optionally registers the context menu for the temporary-save feature.

.PARAMETER MakeLinks
Create links from `home` into $env:USERPROFILE (default: on).

.PARAMETER RegisterContextMenu
Registers the temporary-save right-click menu for the current user.

.PARAMETER TemporaryFolderPath
Folder used by the temporary-save menu (default: TargetRoot/一時保存フォルダ).

.PARAMETER SourceRoot
Path to the `home` folder (default: repo/home).

.PARAMETER TargetRoot
Destination root (default: $env:USERPROFILE).

.PARAMETER Overwrite
Overwrite existing items at the destination (can use -BackupRoot to move them).

.PARAMETER BackupRoot
If provided with -Overwrite, existing items are moved here with timestamp.

.EXAMPLE
./Install-WinTools.ps1 -RegisterContextMenu -Verbose

.EXAMPLE
./Install-WinTools.ps1 -Overwrite -BackupRoot "$env:USERPROFILE\\win-tools_backup"
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [switch]$MakeLinks = $true,
  [switch]$RegisterContextMenu,
  [switch]$RegisterExcelReadOnly,
  [string]$SourceRoot,
  [string]$TargetRoot = $env:USERPROFILE,
  [switch]$Overwrite,
  [string]$BackupRoot,
  [string]$TemporaryFolderPath = (Join-Path $TargetRoot '一時保存フォルダ')
)

$scriptDir = Split-Path -Parent $PSCommandPath
$repoRoot  = Split-Path -Parent $scriptDir
if (-not $SourceRoot) { $SourceRoot = Join-Path $repoRoot 'home' }

# Ensure linking script is available
$linkScript = Join-Path $scriptDir 'New-HomeLinks.ps1'
if (-not (Test-Path -LiteralPath $linkScript)) {
  throw "Missing script: $linkScript"
}

if ($MakeLinks) {
  & $linkScript -SourceRoot $SourceRoot -TargetRoot $TargetRoot -Overwrite:$Overwrite -BackupRoot $BackupRoot -Verbose:$VerbosePreference
}

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

Write-Host "Install complete."
