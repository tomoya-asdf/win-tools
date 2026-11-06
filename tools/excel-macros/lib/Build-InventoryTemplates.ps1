<#
.SYNOPSIS
VBAソースからファイル一覧・シート一覧のExcelテンプレートを生成します。
.DESCRIPTION
専用Excelを非表示で使用します。VBAアクセス設定や信頼済みフォルダは変更しません。
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [ValidateSet('All','File','Sheet')][string]$Kind = 'All',
  [string]$OutputDirectory = (Split-Path -Parent $PSScriptRoot),
  [switch]$Force
)
$directory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw '出力先フォルダがありません。' }
$kinds = if ($Kind -eq 'All') { @('File','Sheet') } elseif ($Kind -eq 'File') { @('File') } else { @('Sheet') }
foreach ($kind in $kinds) {
  $moduleName = $kind + 'Inventory'
  $output = Join-Path $directory ($moduleName + '.xltm')
  $title = if ($kind -eq 'File') { 'フォルダ内のファイル一覧' } else { 'Excelファイルのシート一覧' }
  if ((Test-Path -LiteralPath $output) -and -not $Force) { throw "既存テンプレートの置き換えには-Forceが必要です: $output" }
  if (-not $PSCmdlet.ShouldProcess($output, 'Build Excel inventory template')) { continue }
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
    $project.Name = 'WinTools' + $moduleName
    $common = $project.VBComponents.Add(1)
    $common.Name = 'InventoryCommon'
    $common.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'InventoryCommon.bas')))
    $component = $project.VBComponents.Add(1)
    $component.Name = $moduleName
    $component.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot ($moduleName + '.bas'))))
    $events = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ThisWorkbook.cls')).Replace('FileInventory.StartFileInventory', $moduleName + '.Start' + $moduleName)
    $project.VBComponents.Item($book.CodeName).CodeModule.AddFromString($events)
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = '設定'
    $sheet.Cells.Font.Size = 11
    $sheet.Columns.Item('A').ColumnWidth = 48
    $sheet.Columns.Item('B').ColumnWidth = 65
    [void]$sheet.Range('A1:B1').Merge()
    $sheet.Range('A1').Value2 = $title
    $sheet.Range('A1:B1').Font.Size = 18
    $sheet.Range('A1:B1').Font.Bold = $true
    $sheet.Range('A1:B1').Font.Color = 16777215
    $sheet.Range('A1:B1').Interior.Color = 0x00794E1F
    [void]$sheet.Range('A2:B2').Merge()
    $sheet.Range('A2').Value2 = '起動時に対象を選択し、検索条件を入力します。空欄はすべて対象です。'
    $sheet.Range('B4:B7').NumberFormat = '@'
    $sheet.Range('B4:B7').Interior.Color = 0x00FAF1E5
    $labels = if ($kind -eq 'File') { @('拡張子の正規表現（ドットなし）', 'ファイル名の正規表現', '相対フォルダの正規表現（ルートは .）', '除外フォルダの正規表現') } else { @('ブック名の正規表現', '相対フォルダの正規表現（ルートは .）', '除外フォルダの正規表現', 'シート名の正規表現') }
    for ($i = 0; $i -lt $labels.Count; $i++) { $sheet.Cells.Item($i + 4, 1).Value2 = $labels[$i] }
    $sheet.Range('A9').Value2 = '空欄は全件。正規表現は大文字小文字を区別しません。'
    $sheet.Range('A11').Value2 = if ($kind -eq 'File') { '拡張子の例' } else { 'ブック名の例' }
    $sheet.Range('B11').NumberFormat = '@'
    $sheet.Range('B11').Value2 = if ($kind -eq 'File') { '^(xlsx|xlsm|csv)$' } else { '^売上.*\.(xlsx|xlsm)$' }
    $sheet.Range('A12').Value2 = '相対フォルダの例'
    $sheet.Range('B12').NumberFormat = '@'
    $sheet.Range('B12').Value2 = '(^|\\)請求(\\|$)'
    $sheet.Range('A13').Value2 = '除外フォルダの例'
    $sheet.Range('B13').NumberFormat = '@'
    $sheet.Range('B13').Value2 = '(^|\\)(\.git|node_modules)(\\|$)'
    [void]$sheet.Range('A15:B15').Merge()
    $sheet.Range('A15').Value2 = '結果は別ブックに出力し、この実行用ブックは完了後に保存せず閉じます。'
    [void]$sheet.Range('A17:B17').Merge()
    $sheet.Range('A17').Value2 = '対象収集後に進捗率を表示します。読み取り失敗は［探索エラー］に残します。'
    $button = $sheet.Shapes.AddShape(5, 20, 370, 330, 44)
    $button.Name = 'StartInventory'
    $button.TextFrame2.TextRange.Text = '対象を選んで一覧を作成'
    $button.TextFrame2.TextRange.Font.Size = 14
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    $button.Fill.ForeColor.RGB = 0x00794E1F
    $button.OnAction = $moduleName + '.Start' + $moduleName
    [void]$sheet.Activate()
    [void]$sheet.Range('B4').Select()
    $book.SaveAs($output, 53)
  } finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($item in @($component, $common, $button, $sheet, $project, $book, $excel)) {
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

}
