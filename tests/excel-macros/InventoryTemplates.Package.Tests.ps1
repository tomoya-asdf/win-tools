$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$templatePaths = @('FileInventory.xltm','SheetInventory.xltm') | ForEach-Object { Join-Path $repoRoot ('tools\excel-macros\'+$_) }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

function Read-InventoryPackagePart {
  param([string]$Workbook, [string]$Part)
  $archive = [IO.Compression.ZipFile]::OpenRead($Workbook)
  try {
    $entry = $archive.GetEntry($Part)
    if (-not $entry) { throw "Missing workbook part: $Part" }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally { $archive.Dispose() }
}

Describe 'Distributable inventory templates' {
  It 'contains VBA in both real macro-enabled template packages' {
    foreach ($path in $templatePaths) {
      Read-InventoryPackagePart $path '[Content_Types].xml' | Should Match 'application/vnd.ms-excel.template.macroEnabled.main\+xml'
      $archive = [IO.Compression.ZipFile]::OpenRead($path)
      try {
        $entry = $archive.GetEntry('xl/vbaProject.bin')
        ($null -ne $entry) | Should Be $true
        $stream = $entry.Open()
        try {
          $header = New-Object byte[] 8
          $null = $stream.Read($header, 0, $header.Length)
          [BitConverter]::ToString($header) | Should Be 'D0-CF-11-E0-A1-B1-1A-E1'
        } finally { $stream.Dispose() }
      } finally { $archive.Dispose() }
    }
  }
  It 'uses neutral creator metadata and omits local save directories' {
    foreach ($path in $templatePaths) {
      $core = [xml](Read-InventoryPackagePart $path 'docProps/core.xml')
      foreach ($node in $core.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) { $node.InnerText | Should Be 'win-tools' }
      $workbook = [xml](Read-InventoryPackagePart $path 'xl/workbook.xml')
      $workbook.SelectNodes('//*[local-name()="absPath"]').Count | Should Be 0
    }
  }
  It 'automatically lists both templates in the launcher and removes the old workbook' {
    $items = @(& (Join-Path $repoRoot 'Start-WinTools.ps1') -List)
    foreach ($path in $templatePaths) {
      $macro = @($items | Where-Object { $_.Kind -eq 'ExcelMacro' -and $_.Path -eq $path })
      $macro.Count | Should Be 1
      $macro[0].Name | Should Be ('Excel: ' + (Split-Path -Leaf $path))
    }
    Test-Path -LiteralPath (Join-Path $repoRoot 'tools\excel-macros\FileInventory.xlsm') | Should Be $false
  }
  It 'does not contain user paths, machine names or email addresses in package parts' {
    foreach ($path in $templatePaths) {
      $archive = [IO.Compression.ZipFile]::OpenRead($path)
      try {
        foreach ($entry in $archive.Entries) {
          $stream = $entry.Open(); $memory = New-Object IO.MemoryStream
          try {
            $stream.CopyTo($memory); $bytes = $memory.ToArray()
            foreach ($encoding in @([Text.Encoding]::UTF8, [Text.Encoding]::Unicode)) {
              $text = $encoding.GetString($bytes)
              ($text -match '(?i)[A-Z]:\\Users\\|[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}') | Should Be $false
              if ($env:COMPUTERNAME) { $text.Contains($env:COMPUTERNAME) | Should Be $false }
            }
          } finally { $stream.Dispose(); $memory.Dispose() }
        }
      } finally { $archive.Dispose() }
    }
  }
}
