Option Explicit

Private m_Busy As Boolean
Private m_LastError As String
Private m_Report As Workbook
Private m_Sheets As Worksheet
Private m_Errors As Worksheet
Private m_Info As Worksheet
Private m_Count As Long
Private m_ErrorCount As Long
Private m_Total As Long
Private m_Completed As Long

Public Sub StartSheetInventory()
    Dim mode As Variant, picker As Object, root As String, report As Workbook, settings As Worksheet
    Dim filePattern As String, folderPattern As String, excludedPattern As String, sheetPattern As String
    On Error GoTo Failed
    If m_Busy Then Exit Sub
    mode = Application.InputBox("1: フォルダ内のExcelファイル（子フォルダも探索）" & vbCrLf & _
        "2: Excelファイルを1つ選択", "シート一覧の対象", 1, Type:=1)
    If VarType(mode) = vbBoolean Then GoTo CloseTool
    If mode = 1 Then
        Set picker = Application.FileDialog(4)
        picker.Title = "Excelファイルを探索するフォルダ"
    ElseIf mode = 2 Then
        Set picker = Application.FileDialog(3)
        picker.Title = "シート一覧を取得するExcelファイル"
        picker.Filters.Clear
        picker.Filters.Add "Excelブック", "*.xlsx;*.xlsm;*.xlsb;*.xls;*.xltx;*.xltm;*.xlt;*.xlam;*.xla"
    Else
        MsgBox "1または2を入力してください。", vbExclamation
        GoTo CloseTool
    End If
    picker.AllowMultiSelect = False
    If picker.Show <> -1 Then GoTo CloseTool
    root = picker.SelectedItems(1)
    Set settings = ThisWorkbook.Worksheets("設定")
    If Not InventoryCommon.AskPattern("ブック名の正規表現。空欄はすべて。", CStr(settings.Range("B4").Value2), filePattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("相対フォルダの正規表現。ルートは .、空欄はすべて。", CStr(settings.Range("B5").Value2), folderPattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("除外フォルダの正規表現。空欄は除外なし。", CStr(settings.Range("B6").Value2), excludedPattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("シート名の正規表現。空欄はすべて。", CStr(settings.Range("B7").Value2), sheetPattern) Then GoTo CloseTool
    Set report = BuildSheetInventory(root, filePattern, folderPattern, excludedPattern, sheetPattern)
    If Len(m_LastError) > 0 Then MsgBox m_LastError, vbExclamation, "シート一覧"
    If Not report Is Nothing Then
        report.Activate
        report.Worksheets("シート一覧").Activate
    End If
CloseTool:
    InventoryCommon.CloseInventoryTool report
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "シート一覧"
    Resume CloseTool
End Sub

Public Function LastInventoryError() As String
    LastInventoryError = m_LastError
End Function

Public Function BuildSheetInventory(ByVal targetPath As String, Optional ByVal filePattern As String = "", _
    Optional ByVal folderPattern As String = "", Optional ByVal excludedPattern As String = "", _
    Optional ByVal sheetPattern As String = "") As Workbook
    Dim fs As Object, extensionRegex As Object, fileRegex As Object, folderRegex As Object, excludedRegex As Object, sheetRegex As Object
    Dim candidates As Collection, issues As Collection, entry As Variant, issue As Variant, relative As String
    Dim examined As Long, skipped As Long, excluded As Long, status As String, number As Long
    Dim oldEvents As Boolean, oldUpdating As Boolean, oldDisplayStatus As Boolean, oldAlerts As Boolean
    Dim oldSecurity As Long, oldCancel As Long, oldStatus As Variant, savedState As Boolean
    If m_Busy Then Exit Function
    On Error GoTo Failed
    m_Busy = True
    m_LastError = vbNullString
    InventoryCommon.ResetProgress
    m_Count = 0: m_ErrorCount = 0: m_Total = 0: m_Completed = 0
    Set m_Report = Nothing
    Set extensionRegex = InventoryCommon.MakeRegex("^(xlsx|xlsm|xlsb|xls|xltx|xltm|xlt|xlam|xla)$", "拡張子")
    Set fileRegex = InventoryCommon.MakeRegex(filePattern, "ブック名")
    Set folderRegex = InventoryCommon.MakeRegex(folderPattern, "相対フォルダ")
    Set excludedRegex = InventoryCommon.MakeRegex(excludedPattern, "除外フォルダ")
    Set sheetRegex = InventoryCommon.MakeRegex(sheetPattern, "シート名")
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(targetPath) And Not fs.FolderExists(targetPath) Then Err.Raise vbObjectError + 702, , "対象ファイル・フォルダが見つかりません。"
    If fs.FileExists(targetPath) Then
        If Not extensionRegex.Test(fs.GetExtensionName(targetPath)) Then Err.Raise vbObjectError + 705, , "対応するExcelファイルを指定してください。"
    End If
    oldEvents = Application.EnableEvents
    oldUpdating = Application.ScreenUpdating
    oldDisplayStatus = Application.DisplayStatusBar
    oldAlerts = Application.DisplayAlerts
    oldSecurity = Application.AutomationSecurity
    oldCancel = Application.EnableCancelKey
    oldStatus = Application.StatusBar
    savedState = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = False
    Application.AutomationSecurity = 3
    Application.EnableCancelKey = xlErrorHandler
    InitializeReport
    If fs.FolderExists(targetPath) Then
        Set candidates = InventoryCommon.CollectFiles(targetPath, extensionRegex, fileRegex, folderRegex, excludedRegex, _
            issues, examined, skipped, excluded)
        For Each issue In issues
            AddError CStr(issue(0)), CStr(issue(1))
        Next issue
    Else
        If Not extensionRegex.Test(fs.GetExtensionName(targetPath)) Then Err.Raise vbObjectError + 705, , "対応するExcelファイルを指定してください。"
        Set candidates = New Collection
        examined = 1
        If (fs.GetFile(targetPath).Attributes And 1024) <> 0 Then
            skipped = 1
        ElseIf InventoryCommon.Matches(fileRegex, fs.GetFileName(targetPath)) And InventoryCommon.Matches(folderRegex, ".") Then
            If excludedRegex Is Nothing Then
                candidates.Add Array(fs.GetFile(targetPath).Path, ".")
            ElseIf Not excludedRegex.Test(".") Then
                candidates.Add Array(fs.GetFile(targetPath).Path, ".")
            Else
                excluded = 1
            End If
        End If
    End If
    m_Total = candidates.Count
    InventoryCommon.ShowProgress 0, m_Total, "シート一覧作成中"
    For Each entry In candidates
        ReadWorkbook CStr(entry(0)), CStr(entry(1)), sheetRegex
        m_Completed = m_Completed + 1
        InventoryCommon.ShowProgress m_Completed, m_Total, "シート一覧作成中: " & CStr(entry(0))
    Next entry
    status = "完了"
    GoTo Finish
Failed:
    number = Err.Number
    If number = 18 Then
        m_LastError = "処理を中断しました。": status = "中断"
    Else
        m_LastError = Err.Description: status = "エラー（取得済みの結果）"
    End If
Finish:
    On Error Resume Next
    If Not m_Report Is Nothing Then
        WriteSummary targetPath, filePattern, folderPattern, excludedPattern, sheetPattern, status, skipped, excluded
        FormatReport
        Set BuildSheetInventory = m_Report
    End If
    If savedState Then
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldUpdating
        Application.DisplayStatusBar = oldDisplayStatus
        Application.DisplayAlerts = oldAlerts
        Application.AutomationSecurity = oldSecurity
        Application.EnableCancelKey = oldCancel
        Application.StatusBar = oldStatus
    End If
    m_Busy = False
End Function

Private Sub InitializeReport()
    Dim headers As Variant, i As Long
    Set m_Report = Application.Workbooks.Add(xlWBATWorksheet)
    Set m_Sheets = m_Report.Worksheets(1)
    m_Sheets.Name = "シート一覧"
    Set m_Errors = m_Report.Worksheets.Add(After:=m_Sheets)
    m_Errors.Name = "探索エラー"
    Set m_Info = m_Report.Worksheets.Add(After:=m_Errors)
    m_Info.Name = "実行情報"
    m_Sheets.Columns("B:D").NumberFormat = "@"
    m_Sheets.Columns("F:I").NumberFormat = "@"
    m_Errors.Columns("A:B").NumberFormat = "@"
    m_Info.Columns("A:B").NumberFormat = "@"
    headers = Array("No", "ブック名", "相対パス", "フルパス", "シート順", "シート名", "種類", "表示状態", "相対フォルダ")
    For i = 0 To UBound(headers)
        m_Sheets.Cells(1, i + 1).Value2 = headers(i)
    Next i
    m_Errors.Cells(1, 1).Value2 = "対象パス": m_Errors.Cells(1, 2).Value2 = "内容"
    m_Info.Cells(1, 1).Value2 = "項目": m_Info.Cells(1, 2).Value2 = "値"
End Sub

Private Sub ReadWorkbook(ByVal path As String, ByVal relative As String, ByVal sheetRegex As Object)
    Dim source As Workbook, existing As Workbook, owned As Boolean, sheet As Object
    Dim kind As String, visible As String, relativePath As String, number As Long, message As String
    On Error GoTo Failed
    For Each existing In Application.Workbooks
        If StrComp(existing.FullName, path, vbTextCompare) = 0 Then Set source = existing: Exit For
    Next existing
    If source Is Nothing Then
        Set source = Application.Workbooks.Open(FileName:=path, UpdateLinks:=0, ReadOnly:=True, _
            Password:="", WriteResPassword:="", IgnoreReadOnlyRecommended:=True, Editable:=True, Notify:=False, AddToMru:=False)
        owned = True
    End If
    For Each sheet In source.Sheets
        If InventoryCommon.Matches(sheetRegex, CStr(sheet.Name)) Then
            If m_Count >= m_Sheets.Rows.Count - 1 Then Err.Raise vbObjectError + 703, , "Excelの行数上限に達しました。"
            Select Case TypeName(sheet)
                Case "Worksheet": kind = "ワークシート"
                Case "Chart": kind = "グラフ"
                Case "DialogSheet": kind = "ダイアログ"
                Case Else: kind = TypeName(sheet)
            End Select
            Select Case sheet.Visible
                Case -1: visible = "表示"
                Case 0: visible = "非表示"
                Case 2: visible = "VeryHidden"
                Case Else: visible = CStr(sheet.Visible)
            End Select
            If relative = "." Then relativePath = source.Name Else relativePath = relative & "\" & source.Name
            m_Count = m_Count + 1
            m_Sheets.Cells(m_Count + 1, 1).Resize(1, 9).Value2 = _
                Array(m_Count, source.Name, relativePath, path, sheet.Index, sheet.Name, kind, visible, relative)
        End If
        DoEvents
    Next sheet
    If owned Then source.Close SaveChanges:=False
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If owned And Not source Is Nothing Then source.Close SaveChanges:=False
    On Error GoTo 0
    If number = 18 Or number = vbObjectError + 703 Then Err.Raise number, , message
    AddError path, message
End Sub

Private Sub AddError(ByVal path As String, ByVal message As String)
    m_ErrorCount = m_ErrorCount + 1
    If m_ErrorCount >= m_Errors.Rows.Count Then Exit Sub
    m_Errors.Cells(m_ErrorCount + 1, 1).Value2 = path
    m_Errors.Cells(m_ErrorCount + 1, 2).Value2 = message
End Sub

Private Sub WriteSummary(ByVal root As String, ByVal filePattern As String, ByVal folderPattern As String, _
    ByVal excludedPattern As String, ByVal sheetPattern As String, ByVal status As String, ByVal skipped As Long, ByVal excluded As Long)
    Dim labels As Variant, values As Variant, i As Long
    labels = Array("状態", "対象パス", "ブック名の正規表現", "相対フォルダの正規表現", "除外フォルダの正規表現", _
        "シート名の正規表現", "対象ブック数", "処理済みブック数", "シート件数", "探索エラー件数", _
        "スキップしたリンク数", "除外したフォルダ数", "進捗率", "終了時刻", "中断・エラー内容")
    values = Array(status, root, filePattern, folderPattern, excludedPattern, sheetPattern, CStr(m_Total), CStr(m_Completed), _
        CStr(m_Count), CStr(m_ErrorCount), CStr(skipped), CStr(excluded), CStr(Int(InventoryCommon.LastProgressPercent())) & "%", _
        Format$(Now, "yyyy-mm-dd hh:nn:ss"), m_LastError)
    For i = 0 To UBound(labels)
        m_Info.Cells(i + 2, 1).Value2 = labels(i): m_Info.Cells(i + 2, 2).Value2 = values(i)
    Next i
End Sub

Private Sub FormatReport()
    Dim sheet As Worksheet, lastRow As Long, width As Long, tableRange As Range
    For Each sheet In m_Report.Worksheets
        lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        If sheet.Name = "シート一覧" Then width = 9 Else width = 2
        Set tableRange = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, width))
        With tableRange.Rows(1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
        End With
        tableRange.Columns.AutoFit
        tableRange.AutoFilter
    Next sheet
    m_Sheets.Columns("B:D").ColumnWidth = 32
    m_Sheets.Columns("F:I").ColumnWidth = 24
    m_Errors.Columns("A:B").ColumnWidth = 55
    m_Info.Columns("B").ColumnWidth = 65
End Sub
