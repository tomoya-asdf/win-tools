<#
.SYNOPSIS
実Excelでシート一覧テンプレートを検証します。通常のPesterでは起動しません。
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$template = Join-Path $repoRoot 'tools\excel-macros\SheetInventory.xltm'
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('win-tools-sheets-' + [guid]::NewGuid().ToString('N'))
$excel = $null; $tool = $null; $userBook = $null; $link = $null; $passed = 0
function Assert-Sheets {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) { throw $Message }
}
function Invoke-SheetCase {
  param([string]$Target, [string]$File='', [string]$Folder='', [string]$Exclude='', [string]$Sheet='')
  $report = $excel.Run($prefix + 'SheetInventory.BuildSheetInventory', $Target, $File, $Folder, $Exclude, $Sheet)
  if (-not $report) { throw ('No report: ' + $excel.Run($prefix + 'SheetInventory.LastInventoryError')) }
  return $report
}
function Get-SheetRows {
  param($Report)
  $sheet = $Report.Worksheets.Item('シート一覧')
  $last = $sheet.Cells.Item($sheet.Rows.Count, 1).End(-4162).Row
  for ($i=2; $i -le $last; $i++) {
    [pscustomobject]@{Book=$sheet.Cells.Item($i,2).Value2;Relative=$sheet.Cells.Item($i,3).Value2;Index=$sheet.Cells.Item($i,5).Value2;Name=$sheet.Cells.Item($i,6).Value2;Kind=$sheet.Cells.Item($i,7).Value2;Visible=$sheet.Cells.Item($i,8).Value2}
  }
}
function Get-SheetInfo {
  param($Report,[string]$Name)
  $sheet=$Report.Worksheets.Item('実行情報')
  for($i=2;$i -le 16;$i++){if($sheet.Cells.Item($i,1).Value2 -eq $Name){return $sheet.Cells.Item($i,2).Value2}}
}
function Close-SheetReport {
  param($Report)
  $Report.Close($false)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Report)
}
function Get-SharedFileHash {
  param([string]$Path)
  $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
  $sha=[Security.Cryptography.SHA256]::Create()
  try{return [BitConverter]::ToString($sha.ComputeHash($stream)) -replace '-',''}finally{$sha.Dispose();$stream.Dispose()}
}
try {
  New-Item -ItemType Directory -Path $fixture | Out-Null
  $root=(New-Item -ItemType Directory -Path (Join-Path $fixture '対象')).FullName
  $nested=(New-Item -ItemType Directory -Path (Join-Path $root '子フォルダ')).FullName
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
  $main=Join-Path $root 'main.xlsx'
  $book=$excel.Workbooks.Add(-4167)
  $book.Worksheets.Item(1).Name='通常'
  $hidden=$book.Worksheets.Add(); $hidden.Name='非表示'; $hidden.Visible=0
  $veryHidden=$book.Worksheets.Add(); $veryHidden.Name='強い非表示'; $veryHidden.Visible=2
  $formulaName=$book.Worksheets.Add(); $formulaName.Name='=1+1'
  $chart=$book.Charts.Add(); $chart.Name='グラフ'
  $book.SaveAs($main,51); $book.Close($false)
  $nestedPath=Join-Path $nested 'nested.xlsm'
  $book=$excel.Workbooks.Add(-4167); $book.Worksheets.Item(1).Name='対象シート'
  $marker=Join-Path $fixture 'startup-must-not-run.txt'
  $eventCode='Private Sub Workbook_Open()' + "`r`n" + 'CreateObject("Scripting.FileSystemObject").CreateTextFile("' + $marker + '").Close' + "`r`nEnd Sub"
  $book.VBProject.VBComponents.Item($book.CodeName).CodeModule.AddFromString($eventCode)
  $book.SaveAs($nestedPath,52); $book.Close($false)
  $locked=Join-Path $root 'locked.xlsx'
  $book=$excel.Workbooks.Add(-4167); $book.SaveAs($locked,51,'fixture-only'); $book.Close($false)
  [IO.File]::WriteAllText((Join-Path $root 'broken.xlsx'),'not an Excel workbook')
  [IO.File]::WriteAllText((Join-Path $root 'ordinary.txt'),'not a candidate')
  $outside=(New-Item -ItemType Directory -Path (Join-Path $fixture 'outside')).FullName
  [IO.File]::WriteAllText((Join-Path $outside 'private.xlsx'),'must not inspect')
  $link=Join-Path $root 'linked'; New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
  $before=@{}; foreach($path in @($main,$nestedPath,$locked,$template)){$before[$path]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
  $tool=$excel.Workbooks.Add($template)
  $prefix="'"+$tool.Name.Replace("'","''")+"'!"
  Assert-Sheets ($tool.Path -eq '') 'The source template was opened for editing instead of copied.'
  $module=$tool.VBProject.VBComponents.Item('SheetInventory').CodeModule
  $source=[IO.File]::ReadAllText((Join-Path $repoRoot 'tools\excel-macros\lib\SheetInventory.bas')).Trim()
  Assert-Sheets ($module.Lines(1,$module.CountOfLines).Trim() -eq $source) 'Shipped source is outdated.'
  $report=Invoke-SheetCase $main
  $rows=@(Get-SheetRows $report)
  Assert-Sheets ($rows.Count -eq 5) 'Some sheet types were omitted.'
  Assert-Sheets (($rows | Where-Object Name -eq '非表示').Visible -eq '非表示') 'Hidden sheet not identified.'
  Assert-Sheets (($rows | Where-Object Name -eq '強い非表示').Visible -eq 'VeryHidden') 'VeryHidden sheet not identified.'
  Assert-Sheets (($rows | Where-Object Name -eq 'グラフ').Kind -eq 'グラフ') 'Chart sheet type incorrect.'
  for($i=0;$i -lt $rows.Count;$i++){Assert-Sheets ($rows[$i].Index -eq $i+1) 'Sheet order incorrect.'}
  $sheet=$report.Worksheets.Item('シート一覧');for($i=2;$i -le 6;$i++){Assert-Sheets (-not $sheet.Cells.Item($i,6).HasFormula) 'Sheet name became a formula.'}
  Close-SheetReport $report; Write-Output 'PASS: selected file, sheet order, chart and hidden states';$passed++
  $report=Invoke-SheetCase $root
  Assert-Sheets (@(Get-SheetRows $report).Count -eq 6) 'Recursive scan did not include both readable workbooks.'
  Assert-Sheets ((Get-SheetInfo $report '探索エラー件数') -eq '2') 'Password-protected/corrupt workbooks were not recorded.'
  Assert-Sheets ((Get-SheetInfo $report '進捗率') -eq '100%') 'Failed workbooks were not included in completed progress.'
  Assert-Sheets ((Get-SheetInfo $report 'スキップしたリンク数') -eq '1') 'Linked directory was traversed.'
  Assert-Sheets (-not(Test-Path -LiteralPath $marker)) 'An input workbook startup macro ran.'
  Close-SheetReport $report; Write-Output 'PASS: recursion, locked/corrupt inputs, macro suppression and links';$passed++
  $report=Invoke-SheetCase $root '^nested' '子フォルダ' '' '^対象'
  $rows=@(Get-SheetRows $report);Assert-Sheets ($rows.Count -eq 1 -and $rows[0].Name -eq '対象シート') 'Combined regex filters failed.'
  Close-SheetReport $report; Write-Output 'PASS: combined workbook, folder and sheet regexes';$passed++
  $report=Invoke-SheetCase $root '^(main|nested)' '' '子フォルダ'
  Assert-Sheets (@(Get-SheetRows $report).Count -eq 5) 'Excluded folder was not pruned.'
  Close-SheetReport $report; Write-Output 'PASS: excluded folders';$passed++
  $report=Invoke-SheetCase $main '' '' '' '^no-match$'
  Assert-Sheets (@(Get-SheetRows $report).Count -eq 0) 'No-match filtering failed.'
  Close-SheetReport $report; Write-Output 'PASS: no-match report';$passed++
  $invalid=$excel.Run($prefix+'SheetInventory.BuildSheetInventory',$main,'','','','[')
  Assert-Sheets ($null -eq $invalid) 'Invalid regex should not create a report.'
  Write-Output 'PASS: invalid regex handled before reading inputs';$passed++
  $userBook=$excel.Workbooks.Open($main,0,$false);$userBook.Worksheets.Item('通常').Range('A1').Value2='unsaved user edit'
  $oldSecurity=$excel.AutomationSecurity
  $excel.DisplayStatusBar=$false; $excel.StatusBar='fixture-status'
  $report=Invoke-SheetCase $main
  Assert-Sheets (-not $userBook.Saved -and $userBook.Worksheets.Item('通常').Range('A1').Value2 -eq 'unsaved user edit') 'An already open workbook was altered or closed.'
  Assert-Sheets ($excel.AutomationSecurity -eq $oldSecurity -and $excel.DisplayStatusBar -eq $false -and $excel.StatusBar -eq 'fixture-status') 'Application state was not restored.'
  Close-SheetReport $report; Write-Output 'PASS: existing unsaved workbook and application state preserved';$passed++
  $report=Invoke-SheetCase (Join-Path $repoRoot 'tools\excel-macros\FileInventory.xltm')
  Assert-Sheets (@(Get-SheetRows $report).Count -eq 1) 'Template metadata could not be read.'
  Close-SheetReport $report; Write-Output 'PASS: macro-enabled template input';$passed++
  foreach($path in $before.Keys){Assert-Sheets ((Get-SharedFileHash $path) -eq $before[$path]) 'An input or source template was saved.'}
  $excel.EnableEvents=$true
  $blocked=Join-Path $fixture 'must-not-save.xlsm';try{$tool.SaveAs($blocked,52)}catch{}
  Assert-Sheets (-not(Test-Path -LiteralPath $blocked)) 'Tool save was not blocked.'
  $excel.EnableEvents=$false
  Write-Output 'PASS: templates and source files unchanged; accidental save blocked';$passed++
  $report=Invoke-SheetCase $main
  $name=$tool.Name
  try{$excel.Run($prefix+'InventoryCommon.CloseInventoryTool',$report)|Out-Null}catch{$closeError=$_.Exception.Message}
  $tool=$null
  Assert-Sheets (@($excel.Workbooks|Where-Object Name -eq $name).Count -eq 0) ('Tool did not close: '+$closeError)
  Assert-Sheets ($report.Worksheets.Item('シート一覧').Cells.Item(2,2).Value2 -eq 'main.xlsx') 'Result closed with the tool.'
  Assert-Sheets ($userBook.Worksheets.Item('通常').Range('A1').Value2 -eq 'unsaved user edit') 'Another open workbook closed with the tool.'
  Close-SheetReport $report;Write-Output 'PASS: tool closes while result and user workbook remain open';$passed++
  Write-Output "Sheet Excel integration: $passed cases passed."
} finally {
  if($tool){try{$tool.Close($false)}catch{}}
  if($userBook){try{$userBook.Close($false)}catch{}}
  if($excel){$excel.Quit()}
  foreach($item in @($tool,$userBook,$excel)){if($item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  if($link -and (Test-Path -LiteralPath $link)){[IO.Directory]::Delete($link)}
  $resolved=[IO.Path]::GetFullPath($fixture);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
  if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $resolved).StartsWith('win-tools-sheets-')){Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue}
}
