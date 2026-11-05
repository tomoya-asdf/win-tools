<#
.SYNOPSIS
VBAソースからFileInventory.xlsmを生成します（ExcelとVBAプロジェクトへのアクセスが必要）。
.DESCRIPTION
新しいExcelインスタンスを非表示で使用します。起動イベントは実行しません。
Excelの永続的なセキュリティ設定・信頼済みフォルダ・レジストリの変更は行いません。
.EXAMPLE
.\tools\excel-macros\lib\Build-FileInventoryWorkbook.ps1 -Force
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'FileInventory.xlsm'),
  [switch]$Force
)
$output = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($output) -ne '.xlsm') { throw '出力先は.xlsmを指定してください。' }
if ((Test-Path -LiteralPath $output) -and -not $Force) { throw '既存ブックを置き換える場合は-Forceを指定してください。' }
if (-not (Test-Path -LiteralPath (Split-Path -Parent $output) -PathType Container)) { throw '出力先の親フォルダがありません。' }
if (-not $PSCmdlet.ShouldProcess($output, 'Build Excel file inventory macro workbook')) { return }
$excel = $null
$book = $null
$sheet = $null
$component = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Add(-4167)
  try { $project = $book.VBProject; $null = $project.VBComponents.Count }
  catch { throw 'Excelの「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が必要です。設定は自動変更しません。' }
  $project.Name = 'WinToolsFileInventory'
  $component = $project.VBComponents.Add(1)
  $component.Name = 'FileInventory'
  $component.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'FileInventory.bas')))
  $project.VBComponents.Item($book.CodeName).CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ThisWorkbook.cls')))
  $sheet = $book.Worksheets.Item(1)
  $sheet.Name = '設定'
  $sheet.Cells.Font.Size = 11
  $sheet.Columns.Item('A').ColumnWidth = 48
  $sheet.Columns.Item('B').ColumnWidth = 65
  [void]$sheet.Range('A1:B1').Merge()
  $sheet.Range('A1').Value2 = 'フォルダ内のファイル一覧'
  $sheet.Range('A1:B1').Font.Size = 18
  $sheet.Range('A1:B1').Font.Bold = $true
  $sheet.Range('A1:B1').Font.Color = 16777215
  $sheet.Range('A1:B1').Interior.Color = 0x00794E1F
  [void]$sheet.Range('A2:B2').Merge()
  $sheet.Range('A2').Value2 = '起動時にフォルダを選択し、検索条件を入力します。空欄はすべて対象です。'
  $sheet.Range('B4:B7').NumberFormat = '@'
  $sheet.Range('B4:B7').Interior.Color = 0x00FAF1E5
  $labels = @('拡張子の正規表現（ドットなし）', 'ファイル名の正規表現', '相対フォルダの正規表現（ルートは .）', '除外フォルダの正規表現')
  for ($i = 0; $i -lt $labels.Count; $i++) { $sheet.Cells.Item($i + 4, 1).Value2 = $labels[$i] }
  $sheet.Range('A9').Value2 = '空欄は全件。正規表現は大文字小文字を区別しません。'
  $sheet.Range('A11').Value2 = '拡張子の例'
  $sheet.Range('B11').NumberFormat = '@'
  $sheet.Range('B11').Value2 = '^(xlsx|xlsm|csv)$'
  $sheet.Range('A12').Value2 = '相対フォルダの例'
  $sheet.Range('B12').NumberFormat = '@'
  $sheet.Range('B12').Value2 = '(^|\\)請求(\\|$)'
  $sheet.Range('A13').Value2 = '除外フォルダの例'
  $sheet.Range('B13').NumberFormat = '@'
  $sheet.Range('B13').Value2 = '(^|\\)(\.git|node_modules)(\\|$)'
  [void]$sheet.Range('A15:B15').Merge()
  $sheet.Range('A15').Value2 = '一覧は別の新しいブックへ出力します。必要に応じて名前を付けて保存してください。'
  [void]$sheet.Range('A17:B17').Merge()
  $sheet.Range('A17').Value2 = '子フォルダも探索します。リンクはたどらず、読み取り失敗は［探索エラー］に残します。'
  $button = $sheet.Shapes.AddShape(5, 20, 370, 330, 44)
  $button.Name = 'StartInventory'
  $button.TextFrame2.TextRange.Text = 'フォルダを選んで一覧を作成'
  $button.TextFrame2.TextRange.Font.Size = 14
  $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
  $button.Fill.ForeColor.RGB = 0x00794E1F
  $button.OnAction = "'" + [IO.Path]::GetFileName($output).Replace("'", "''") + "'!FileInventory.StartFileInventory"
  [void]$sheet.Activate()
  [void]$sheet.Range('B4').Select()
  $book.SaveAs($output, 52)
} finally {
  if ($book) { $book.Close($false) }
  if ($excel) { $excel.Quit() }
  foreach ($item in @($component, $sheet, $book, $excel)) {
    if ($item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) }
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

# Excel may write a user's name and the local save directory when saving.
# Remove these from the distributable workbook without changing the VBA payload.
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::Open($output, [IO.Compression.ZipArchiveMode]::Update)
try {
  foreach ($name in @('docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml')) {
    $entry = $archive.GetEntry($name)
    if (-not $entry) { continue }
    $stream = $entry.Open()
    try { $xml = New-Object Xml.XmlDocument; $xml.Load($stream) } finally { $stream.Dispose() }
    if ($name -eq 'docProps/core.xml') {
      foreach ($node in $xml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) { $node.InnerText = 'win-tools' }
    } elseif ($name -eq 'docProps/app.xml') {
      foreach ($node in $xml.SelectNodes('//*[local-name()="Company" or local-name()="Manager" or local-name()="HyperlinkBase"]')) { $node.InnerText = '' }
    } else {
      foreach ($node in @($xml.SelectNodes('//*[local-name()="absPath"]'))) { [void]$node.ParentNode.RemoveChild($node) }
    }
    $entry.Delete()
    $newEntry = $archive.CreateEntry($name)
    $stream = $newEntry.Open()
    try { $xml.Save($stream) } finally { $stream.Dispose() }
  }
} finally { $archive.Dispose() }
Get-Item -LiteralPath $output
