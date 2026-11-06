$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'tools\launcher\lib\WinToolsHistory.psm1') -Force

Describe 'Execution history persistence' {
  BeforeEach {
    $historyFolder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $record = [pscustomobject]@{
      SchemaVersion = 1; Id = 'fixture'; StartTime = [DateTimeOffset]::Now.ToString('o')
      EndTime = [DateTimeOffset]::Now.ToString('o'); Name = '比較 [1]'; Kind = 'FolderCompare'
      ToolPath = 'C:\作業\compare.ps1'; Status = 'Succeeded'; DurationMs = 123; Error = $null
    }
  }
  It 'returns no history before the first run without creating directories' {
    @(Get-WinToolsHistory -Directory $historyFolder).Count | Should Be 0
    Test-Path -LiteralPath $historyFolder | Should Be $false
  }
  It 'appends separate Unicode records and returns the latest records first' {
    Write-WinToolsHistory -Record $record -Directory $historyFolder
    $record.Id = 'second'; $record.Error = "日本語のエラー`r`n次の行"
    Write-WinToolsHistory -Record $record -Directory $historyFolder
    $rows = @(Get-WinToolsHistory -Directory $historyFolder)
    $rows.Count | Should Be 2
    $rows[0].Id | Should Be 'second'
    $rows[1].Name | Should Be '比較 [1]'
    $rows[0].Error | Should Be "日本語のエラー`r`n次の行"
    @(Get-Content -LiteralPath (Get-ChildItem -LiteralPath $historyFolder)[0].FullName).Count | Should Be 2
    @(Get-WinToolsHistory -Directory $historyFolder -Last 1).Count | Should Be 1
  }
  It 'reads across monthly files and skips damaged records without hiding valid history' {
    Write-WinToolsHistory -Record $record -Directory $historyFolder
    $older = $record | ConvertTo-Json -Compress
    [IO.File]::WriteAllText((Join-Path $historyFolder '2000-01.history.jsonl'), $older, [Text.UTF8Encoding]::new($false))
    $current = (Get-ChildItem -LiteralPath $historyFolder | Sort-Object Name -Descending)[0].FullName
    [IO.File]::AppendAllText($current, "broken JSON`r`n", [Text.UTF8Encoding]::new($false))
    $warnings = @()
    $rows = @(Get-WinToolsHistory -Directory $historyFolder -Last 2 -WarningVariable warnings -WarningAction SilentlyContinue)
    $rows.Count | Should Be 2
    $warnings.Count | Should Be 1
  }
  It 'exposes history as objects for filtering and export through its script entry point' {
    Write-WinToolsHistory -Record $record -Directory $historyFolder
    $rows = @(& (Join-Path $repoRoot 'tools\launcher\Get-WinToolsHistory.ps1') -Directory $historyFolder -Last 1)
    $rows.Count | Should Be 1
    $rows[0].Kind | Should Be 'FolderCompare'
  }
}
