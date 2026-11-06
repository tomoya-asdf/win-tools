<#
.SYNOPSIS
実際のExcelでファイル一覧マクロを検証します。通常のPesterからは起動しません。
.DESCRIPTION
非表示の専用Excelインスタンスで、ローカルの既知のマクロだけを実行します。
起動イベントを無効にし、フォルダ選択ダイアログは開きません。テスト結果は保存しません。
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$workbookPath = Join-Path $repoRoot 'tools\excel-macros\FileInventory.xltm'
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('win-tools-inventory-' + [guid]::NewGuid().ToString('N'))
$excel = $null
$book = $null
$link = $null
$passed = 0

function Assert-Inventory {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) { throw $Message }
}
function Invoke-InventoryCase {
  param([string]$Root, [string]$Extension = '', [string]$FileName = '', [string]$Folder = '', [string]$Exclude = '')
  return $excel.Run(($macroPrefix + 'FileInventory.BuildFileInventory'), $Root, $Extension, $FileName, $Folder, $Exclude)
}
function Get-InventoryPaths {
  param($Report)
  $sheet = $Report.Worksheets.Item('ファイル一覧')
  $last = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
  for ($i = 2; $i -le $last; $i++) { [string]$sheet.Cells.Item($i, 5).Value2 }
}
function Assert-InventoryPaths {
  param([string]$Name, $Report, [string[]]$Expected)
  Assert-Inventory ($null -ne $Report) "$Name : report was not created."
  $actual = @(Get-InventoryPaths -Report $Report | Sort-Object)
  $expectedSorted = @($Expected | Sort-Object)
  Assert-Inventory (($actual -join "`n") -eq ($expectedSorted -join "`n")) "$Name : unexpected file list."
  Assert-Inventory ($Report.Worksheets.Item('実行情報').Cells.Item(2, 2).Value2 -eq '完了') "$Name : scan did not finish."
  $Report.Close($false)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Report)
  Write-Output "PASS: $Name"
}

try {
  New-Item -ItemType Directory -Path $fixture | Out-Null
  $root = (New-Item -ItemType Directory -Path (Join-Path $fixture '選択フォルダ')).FullName
  foreach ($relative in @('plain.txt','=1+1.txt','README','A\請求\invoice.CSV','A\請求\深い\detail.txt','A\other\keep.pdf','skip\hidden.txt')) {
    $path = Join-Path $root $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
    [IO.File]::WriteAllText($path, 'fixture')
  }
  $outside = (New-Item -ItemType Directory -Path (Join-Path $fixture 'outside')).FullName
  [IO.File]::WriteAllText((Join-Path $outside 'outside.txt'), 'must not follow')
  $link = Join-Path $root 'linked'
  New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
  $beforeHash = (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $book = $excel.Workbooks.Add($workbookPath)
$macroPrefix = "'" + $book.Name.Replace("'", "''") + "'!"
  $source = $book.VBProject.VBComponents.Item($book.CodeName).CodeModule
  Assert-Inventory ($source.Lines(1, $source.CountOfLines).Contains('Private Sub Workbook_Open()')) 'Workbook open event is missing.'
  Assert-Inventory ($source.Lines(1, $source.CountOfLines).Contains('FileInventory.StartFileInventory')) 'Workbook open event is not wired to the picker.'
  $module = $book.VBProject.VBComponents.Item('FileInventory').CodeModule
  $expectedSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\excel-macros\lib\FileInventory.bas')).Trim()
  Assert-Inventory ($module.Lines(1, $module.CountOfLines).Trim() -eq $expectedSource) 'The shipped VBA differs from its reviewable source.'
  $expected = @('plain.txt','=1+1.txt','README','A\請求\invoice.CSV','A\請求\深い\detail.txt','A\other\keep.pdf','skip\hidden.txt')
  $report = Invoke-InventoryCase -Root $root
  Assert-Inventory ($report.Worksheets.Item('実行情報').Cells.Item(11, 2).Value2 -eq '1') 'The junction was not counted as skipped.'
  $files = $report.Worksheets.Item('ファイル一覧')
  $last = $files.Cells.Item($files.Rows.Count, 1).End(-4162).Row
  for ($i = 2; $i -le $last; $i++) {
    Assert-Inventory ($files.Cells.Item($i, 7).Value2 -eq 7) 'The byte size is incorrect.'
    Assert-Inventory ($files.Cells.Item($i, 8).Value2 -gt 0) 'The modified date was not written.'
    if ($files.Cells.Item($i, 2).Value2 -eq '=1+1.txt') {
      Assert-Inventory (-not $files.Cells.Item($i, 2).HasFormula) 'A filename was evaluated as an Excel formula.'
    }
  }
  Assert-InventoryPaths 'recursive enumeration and link exclusion' $report $expected; $passed++
  Assert-InventoryPaths 'extension regex is case insensitive' (Invoke-InventoryCase $root '^(csv)$') @('A\請求\invoice.CSV'); $passed++
  Assert-InventoryPaths 'file name regex' (Invoke-InventoryCase $root '' '^invoice\.') @('A\請求\invoice.CSV'); $passed++
  Assert-InventoryPaths 'include folder does not prune nonmatching parents' (Invoke-InventoryCase $root '' '' '(^|\\)請求(\\|$)') @('A\請求\invoice.CSV','A\請求\深い\detail.txt'); $passed++
  Assert-InventoryPaths 'excluded folder prunes its descendants' (Invoke-InventoryCase $root '' '' '' '(^|\\)skip(\\|$)') @('plain.txt','=1+1.txt','README','A\請求\invoice.CSV','A\請求\深い\detail.txt','A\other\keep.pdf'); $passed++
  Assert-InventoryPaths 'selected root uses a dot as its relative folder' (Invoke-InventoryCase $root '' '' '^\.$') @('plain.txt','=1+1.txt','README'); $passed++
  Assert-InventoryPaths 'files without extensions' (Invoke-InventoryCase $root '^$') @('README'); $passed++
  Assert-InventoryPaths 'combined regular expressions use AND conditions' (Invoke-InventoryCase $root '^csv$' '^invoice' '請求' 'skip') @('A\請求\invoice.CSV'); $passed++
  Assert-InventoryPaths 'zero matches retains result headers' (Invoke-InventoryCase $root '^no-match$') @(); $passed++
  $count = $excel.Workbooks.Count
  $invalid = Invoke-InventoryCase $root '['
  Assert-Inventory ($null -eq $invalid) 'An invalid regex should not create a report.'
  Assert-Inventory ($excel.Workbooks.Count -eq $count) 'An invalid regex left an extra workbook open.'
  Assert-Inventory ($excel.Run(($macroPrefix + 'FileInventory.LastInventoryError')).Contains('拡張子')) 'The invalid regex was not explained.'
  Write-Output 'PASS: invalid regex is handled without a modal error'; $passed++
  $invalid = Invoke-InventoryCase (Join-Path $root 'missing')
  Assert-Inventory ($null -eq $invalid) 'A missing root should not create a report.'
  Write-Output 'PASS: missing root is handled'; $passed++
  $large = (New-Item -ItemType Directory -Path (Join-Path $fixture 'batch')).FullName
  for ($i = 0; $i -lt 1105; $i++) { [IO.File]::WriteAllText((Join-Path $large ('file-{0:D4}.txt' -f $i)), 'x') }
  $report = Invoke-InventoryCase $large
  Assert-Inventory (@(Get-InventoryPaths $report).Count -eq 1105) 'Batch writes lost or duplicated rows.'
  Assert-Inventory ($report.Worksheets.Item('ファイル一覧').Cells.Item(1106, 1).Value2 -eq 1105) 'Sorted row numbering is incorrect.'
  $report.Close($false)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($report)
  Write-Output 'PASS: buffer boundary and sort with 1105 files'; $passed++
  Assert-Inventory ($excel.EnableEvents -eq $false) 'Macro did not restore EnableEvents.'
  Assert-Inventory ($book.Worksheets.Count -eq 1) 'The tool workbook was used to store scan results.'
  Assert-Inventory ($null -eq $book.Worksheets.Item('設定').Range('B4').Value2) 'Test search criteria leaked into the template.'
  Assert-Inventory ($beforeHash -eq (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash) 'The distributed workbook was changed by testing.'
  Write-Output 'PASS: tool template and application event state preserved'; $passed++
  Assert-Inventory ($excel.Run($macroPrefix + 'InventoryCommon.LastProgressPercent') -eq 100) 'Final percentage was not 100.'
  $excel.Run($macroPrefix + 'InventoryCommon.ShowProgress', 1, 4, 'fixture-progress') | Out-Null
  Assert-Inventory ($excel.StatusBar.Contains('25% (1/4)')) 'Intermediate progress did not show the expected percentage.'
  Write-Output 'PASS: intermediate and final percentages'; $passed++
  $excel.EnableEvents = $true
  $blockedPath = Join-Path $fixture 'must-not-save.xlsm'
  try { $book.SaveAs($blockedPath, 52) } catch { }
  Assert-Inventory (-not (Test-Path -LiteralPath $blockedPath)) 'The tool allowed an accidental save.'
  $excel.EnableEvents = $false
  Write-Output 'PASS: accidental tool save is blocked'; $passed++
  $report = Invoke-InventoryCase $root '^csv$'
  $toolName = $book.Name
  try { $excel.Run($macroPrefix + 'InventoryCommon.CloseInventoryTool', $report) | Out-Null }
  catch { $closeError = $_.Exception.Message }
  $book = $null
  Assert-Inventory (@($excel.Workbooks | Where-Object Name -eq $toolName).Count -eq 0) ("The execution workbook did not close: " + $closeError)
  Assert-Inventory ($report.Worksheets.Item('ファイル一覧').Cells.Item(2, 2).Value2 -eq 'invoice.CSV') 'Closing the tool also closed or changed its report.'
  $report.Close($false)
  Write-Output 'PASS: execution workbook closes while result stays open'; $passed++
  Write-Output "Excel integration: $passed cases passed."
} finally {
  if ($book) { try { $book.Close($false) } catch {} }
  if ($excel) { $excel.Quit() }
  foreach ($item in @($book, $excel)) {
    if ($item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) }
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($link -and (Test-Path -LiteralPath $link)) { [IO.Directory]::Delete($link) }
  $resolvedFixture = [IO.Path]::GetFullPath($fixture)
  $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
  if ($resolvedFixture.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $resolvedFixture).StartsWith('win-tools-inventory-')) {
    Remove-Item -LiteralPath $resolvedFixture -Recurse -Force -ErrorAction SilentlyContinue
  }
}
