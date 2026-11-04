<#
.
.SYNOPSIS
Creates links from the repo's `home` items into the user's profile.

.DESCRIPTION
Links each top-level item under `home` to a destination root (default: $env:USERPROFILE).
For directories, creates a Junction (no admin required). For files, tries a SymbolicLink,
falls back to HardLink, then to Copy if needed.

.PARAMETER SourceRoot
Path to the source `home` folder. Defaults to ../home relative to this script.

.PARAMETER TargetRoot
Destination root (defaults to $env:USERPROFILE).

.PARAMETER Overwrite
Remove and recreate destination if it already exists.

.PARAMETER BackupRoot
If set with -Overwrite, moves existing items to this folder instead of deleting.

.EXAMPLE
./New-HomeLinks.ps1 -Verbose

.EXAMPLE
./New-HomeLinks.ps1 -TargetRoot $env:USERPROFILE -Overwrite -BackupRoot "$env:USERPROFILE\\win-tools_backup"
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
  [string]$SourceRoot = (Resolve-Path -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'home')).Path,
  [string]$TargetRoot = $env:USERPROFILE,
  [switch]$Overwrite,
  [string]$BackupRoot
)

function New-LinkForItem {
  param(
    [Parameter(Mandatory=$true)][System.IO.FileSystemInfo]$Source,
    [Parameter(Mandatory=$true)][string]$Destination
  )

  if ($Source.PSIsContainer) {
    # Directory: prefer Junction (no admin needed)
    New-Item -ItemType Junction -Path $Destination -Target $Source.FullName -Force | Out-Null
  } else {
    $created = $false
    try {
      New-Item -ItemType SymbolicLink -Path $Destination -Target $Source.FullName -Force | Out-Null
      $created = $true
    } catch {
      Write-Verbose "SymbolicLink failed for '$($Source.Name)': $($_.Exception.Message)"
    }
    if (-not $created) {
      try {
        New-Item -ItemType HardLink -Path $Destination -Target $Source.FullName -Force | Out-Null
        $created = $true
      } catch {
        Write-Verbose "HardLink failed for '$($Source.Name)': $($_.Exception.Message)"
      }
    }
    if (-not $created) {
      Copy-Item -LiteralPath $Source.FullName -Destination $Destination -Force
      $created = $true
    }
  }
}

function Backup-ExistingItemIfNeeded {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [string]$BackupRoot
  )
  if (-not (Test-Path -LiteralPath $Path)) { return }
  if (-not $BackupRoot) {
    # No backup root specified; remove
    if ((Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue).PSIsContainer) {
      Remove-Item -LiteralPath $Path -Recurse -Force
    } else {
      Remove-Item -LiteralPath $Path -Force
    }
    return
  }

  $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $destDir = Join-Path -Path $BackupRoot -ChildPath $timestamp
  $null = New-Item -ItemType Directory -Path $destDir -Force
  Move-Item -LiteralPath $Path -Destination (Join-Path $destDir (Split-Path -Path $Path -Leaf)) -Force
}

if (-not (Test-Path -LiteralPath $SourceRoot)) {
  throw "SourceRoot not found: $SourceRoot"
}
if (-not (Test-Path -LiteralPath $TargetRoot)) {
  throw "TargetRoot not found: $TargetRoot"
}

$items = Get-ChildItem -LiteralPath $SourceRoot -Force | Where-Object { $_.Name -ne '.' -and $_.Name -ne '..' }
if (-not $items) {
  Write-Warning "No items found under: $SourceRoot"
  return
}

foreach ($it in $items) {
  $dest = Join-Path -Path $TargetRoot -ChildPath $it.Name
  $op = if ($it.PSIsContainer) { 'link (junction)' } else { 'link (symlink/hardlink/copy)' }

  if (Test-Path -LiteralPath $dest) {
    if (-not $Overwrite) {
      Write-Verbose "Exists, skipping: $dest"
      continue
    }
    if ($PSCmdlet.ShouldProcess($dest, "Backup and overwrite")) {
      Backup-ExistingItemIfNeeded -Path $dest -BackupRoot $BackupRoot
    } else { continue }
  }

  $destParent = Split-Path -Path $dest -Parent
  if (-not (Test-Path -LiteralPath $destParent)) { New-Item -ItemType Directory -Path $destParent -Force | Out-Null }

  if ($PSCmdlet.ShouldProcess($dest, "Create $op from '$($it.FullName)'")) {
    try {
      New-LinkForItem -Source $it -Destination $dest
      Write-Host "Linked: $($it.FullName) -> $dest"
    } catch {
      Write-Error "Failed to link '$($it.FullName)' to '$dest': $($_.Exception.Message)"
    }
  }
}
