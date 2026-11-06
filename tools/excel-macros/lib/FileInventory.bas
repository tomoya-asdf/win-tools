Option Explicit

Private Const BUFFER_SIZE As Long = 1000
Private Const COLUMN_COUNT As Long = 8
Private Const REPARSE_POINT As Long = 1024
Private m_Busy As Boolean
Private m_LastError As String
Private m_Report As Workbook
Private m_Files As Worksheet
Private m_Errors As Worksheet
Private m_Info As Worksheet
Private m_Buffer(1 To BUFFER_SIZE, 1 To COLUMN_COUNT) As Variant
Private m_BufferCount As Long
Private m_Written As Long
Private m_Matched As Long
Private m_Examined As Long
Private m_ErrorCount As Long
Private m_SkippedLinks As Long
Private m_ExcludedFolders As Long
Private m_Total As Long
Private m_Completed As Long

Public Sub StartFileInventory()
    Dim picker As Object, settings As Worksheet
    Dim root As String, extensionPattern As String, filePattern As String
    Dim folderPattern As String, excludedPattern As String, report As Workbook
    On Error GoTo Failed
    If m_Busy Then
        MsgBox "現在探索中です。完了するまでお待ちください。", vbInformation
        Exit Sub
    End If
    Set picker = Application.FileDialog(4)
    picker.Title = "ファイル一覧を作成するフォルダを選択"
    picker.AllowMultiSelect = False
    If picker.Show <> -1 Then GoTo CloseTool
    root = picker.SelectedItems(1)
    Set settings = ThisWorkbook.Worksheets("設定")
    If Not InventoryCommon.AskPattern("拡張子（ドットなし）の正規表現。空欄はすべて。" & vbCrLf & _
        "例: ^(xlsx|xlsm|csv)$", CStr(settings.Range("B4").Value2), extensionPattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("ファイル名の正規表現。空欄はすべて。" & vbCrLf & _
        "例: ^売上.*", CStr(settings.Range("B5").Value2), filePattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("相対フォルダの正規表現。ルートは .、空欄はすべて。" & vbCrLf & _
        "例: (^|\\)請求(\\|$)", CStr(settings.Range("B6").Value2), folderPattern) Then GoTo CloseTool
    If Not InventoryCommon.AskPattern("探索しないフォルダの正規表現。空欄は除外なし。" & vbCrLf & _
        "例: (^|\\)(\.git|node_modules)(\\|$)", CStr(settings.Range("B7").Value2), excludedPattern) Then GoTo CloseTool
    Set report = BuildFileInventory(root, extensionPattern, filePattern, folderPattern, excludedPattern)
    If Len(LastInventoryError()) > 0 Then MsgBox LastInventoryError(), vbExclamation, "ファイル一覧"
    If Not report Is Nothing Then
        report.Activate
        report.Worksheets("ファイル一覧").Activate
    End If
CloseTool:
    InventoryCommon.CloseInventoryTool report
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "ファイル一覧"
    Resume CloseTool
End Sub

Public Function LastInventoryError() As String
    LastInventoryError = m_LastError
End Function

' This non-interactive entry point is also used for automated Excel validation.
' Errors are returned via LastInventoryError, never as unattended Excel dialogs.
Public Function BuildFileInventory(ByVal rootPath As String, _
    Optional ByVal extensionPattern As String = "", Optional ByVal filePattern As String = "", _
    Optional ByVal folderPattern As String = "", Optional ByVal excludedPattern As String = "") As Workbook
    Dim fs As Object, extensionRegex As Object, fileRegex As Object
    Dim folderRegex As Object, excludedRegex As Object, pending As Collection
    Dim root As String, prefix As String, folderPath As String, relative As String
    Dim oldUpdating As Boolean, oldEvents As Boolean, oldCancel As Long, oldStatus As Variant
    Dim savedState As Boolean, status As String, errorNumber As Long, oldDisplayStatus As Boolean
    Dim candidates As Collection, issues As Collection, entry As Variant, issue As Variant
    If m_Busy Then
        m_LastError = "現在探索中です。"
        Exit Function
    End If
    On Error GoTo Failed
    m_Busy = True
    InventoryCommon.ResetProgress
    m_Total = 0: m_Completed = 0
    m_LastError = vbNullString
    Set m_Report = Nothing
    Set m_Files = Nothing
    Set m_Errors = Nothing
    Set m_Info = Nothing
    m_BufferCount = 0: m_Written = 0: m_Matched = 0: m_Examined = 0
    m_ErrorCount = 0: m_SkippedLinks = 0: m_ExcludedFolders = 0
    Set extensionRegex = InventoryCommon.MakeRegex(extensionPattern, "拡張子")
    Set fileRegex = InventoryCommon.MakeRegex(filePattern, "ファイル名")
    Set folderRegex = InventoryCommon.MakeRegex(folderPattern, "相対フォルダ")
    Set excludedRegex = InventoryCommon.MakeRegex(excludedPattern, "除外フォルダ")
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(rootPath) Then Err.Raise vbObjectError + 702, , "探索フォルダが見つかりません。"
    root = fs.GetFolder(rootPath).Path
    prefix = root
    If Right$(prefix, 1) <> "\" Then prefix = prefix & "\"
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCancel = Application.EnableCancelKey
    oldStatus = Application.StatusBar
    oldDisplayStatus = Application.DisplayStatusBar
    savedState = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler
    InitializeReport
    Set candidates = InventoryCommon.CollectFiles(root, extensionRegex, fileRegex, folderRegex, excludedRegex, _
        issues, m_Examined, m_SkippedLinks, m_ExcludedFolders)
    For Each issue In issues
        AddError CStr(issue(0)), CStr(issue(1))
    Next issue
    m_Total = candidates.Count
    InventoryCommon.ShowProgress 0, m_Total, "一覧作成中"
    For Each entry In candidates
        AddFile fs, CStr(entry(0)), CStr(entry(1)), extensionRegex, fileRegex, folderRegex
        m_Completed = m_Completed + 1
        InventoryCommon.ShowProgress m_Completed, m_Total, "一覧作成中: " & CStr(entry(0))
    Next entry
    status = "完了"
    GoTo Finish
Failed:
    errorNumber = Err.Number
    If errorNumber = 18 Then
        m_LastError = "探索を中断しました。"
        status = "中断"
    Else
        m_LastError = Err.Description
        status = "エラー（取得済みの結果）"
    End If
Finish:
    On Error Resume Next
    If Not m_Report Is Nothing Then
        FlushRows
        WriteSummary root, extensionPattern, filePattern, folderPattern, excludedPattern, status
        FormatReport
        Set BuildFileInventory = m_Report
    End If
    If savedState Then
        Application.ScreenUpdating = oldUpdating
        Application.EnableEvents = oldEvents
        Application.EnableCancelKey = oldCancel
        Application.StatusBar = oldStatus
        Application.DisplayStatusBar = oldDisplayStatus
    End If
    m_Busy = False
End Function

Private Sub InitializeReport()
    Dim headers As Variant, i As Long
    Set m_Report = Application.Workbooks.Add(xlWBATWorksheet)
    Set m_Files = m_Report.Worksheets(1)
    m_Files.Name = "ファイル一覧"
    Set m_Errors = m_Report.Worksheets.Add(After:=m_Files)
    m_Errors.Name = "探索エラー"
    Set m_Info = m_Report.Worksheets.Add(After:=m_Errors)
    m_Info.Name = "実行情報"
    m_Files.Columns("B:F").NumberFormat = "@"
    m_Files.Columns("A").NumberFormat = "0"
    m_Files.Columns("G").NumberFormat = "0"
    m_Files.Columns("H").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    m_Errors.Columns("A:B").NumberFormat = "@"
    m_Info.Columns("A:B").NumberFormat = "@"
    headers = Array("No", "ファイル名", "拡張子", "相対フォルダ", "相対パス", "フルパス", "サイズ(byte)", "更新日時")
    For i = 0 To UBound(headers)
        m_Files.Cells(1, i + 1).Value2 = headers(i)
    Next i
    m_Errors.Cells(1, 1).Value2 = "対象パス"
    m_Errors.Cells(1, 2).Value2 = "内容"
    m_Info.Cells(1, 1).Value2 = "項目"
    m_Info.Cells(1, 2).Value2 = "値"
End Sub

Private Sub AddFile(ByVal fs As Object, ByVal path As String, ByVal relative As String, _
    ByVal extensionRegex As Object, ByVal fileRegex As Object, ByVal folderRegex As Object)
    Dim extension As String, fileName As String, fullPath As String, number As Long, message As String
    On Error GoTo Failed
    fullPath = path
    Dim item As Object
    Set item = fs.GetFile(path)
    fullPath = item.Path
    If (item.Attributes And REPARSE_POINT) <> 0 Then
        m_SkippedLinks = m_SkippedLinks + 1
        Exit Sub
    End If
    fileName = item.Name
    extension = fs.GetExtensionName(fileName)
    If Not InventoryCommon.Matches(extensionRegex, extension) Then Exit Sub
    If Not InventoryCommon.Matches(fileRegex, fileName) Then Exit Sub
    If Not InventoryCommon.Matches(folderRegex, relative) Then Exit Sub
    If m_Matched >= m_Files.Rows.Count - 1 Then Err.Raise vbObjectError + 703, , "Excelの行数上限に達したため探索を終了しました。"
    ' Read potentially failing metadata before counting this as a result.
    Dim size As Double, modified As Double
    size = CDbl(item.Size): modified = CDbl(item.DateLastModified)
    m_BufferCount = m_BufferCount + 1
    m_Matched = m_Matched + 1
    m_Buffer(m_BufferCount, 1) = m_Matched
    m_Buffer(m_BufferCount, 2) = fileName
    m_Buffer(m_BufferCount, 3) = extension
    m_Buffer(m_BufferCount, 4) = relative
    If relative = "." Then
        m_Buffer(m_BufferCount, 5) = fileName
    Else
        m_Buffer(m_BufferCount, 5) = relative & "\" & fileName
    End If
    m_Buffer(m_BufferCount, 6) = fullPath
    m_Buffer(m_BufferCount, 7) = size
    m_Buffer(m_BufferCount, 8) = modified
    If m_BufferCount = BUFFER_SIZE Then FlushRows
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    If number = 18 Or number = vbObjectError + 703 Then Err.Raise number, , message
    AddError fullPath, message
End Sub

Private Sub FlushRows()
    Dim rows() As Variant, r As Long, c As Long
    If m_BufferCount = 0 Then Exit Sub
    ReDim rows(1 To m_BufferCount, 1 To COLUMN_COUNT)
    For r = 1 To m_BufferCount
        For c = 1 To COLUMN_COUNT
            rows(r, c) = m_Buffer(r, c)
        Next c
    Next r
    m_Files.Cells(m_Written + 2, 1).Resize(m_BufferCount, COLUMN_COUNT).Value2 = rows
    m_Written = m_Written + m_BufferCount
    m_BufferCount = 0
End Sub

Private Sub AddError(ByVal path As String, ByVal message As String)
    m_ErrorCount = m_ErrorCount + 1
    If m_ErrorCount >= m_Errors.Rows.Count Then Exit Sub
    m_Errors.Cells(m_ErrorCount + 1, 1).Value2 = path
    m_Errors.Cells(m_ErrorCount + 1, 2).Value2 = message
End Sub

Private Sub WriteSummary(ByVal root As String, ByVal extensionPattern As String, ByVal filePattern As String, _
    ByVal folderPattern As String, ByVal excludedPattern As String, ByVal status As String)
    Dim labels As Variant, values As Variant, i As Long
    labels = Array("状態", "探索ルート", "拡張子の正規表現", "ファイル名の正規表現", "相対フォルダの正規表現", _
        "除外フォルダの正規表現", "一致件数", "確認したファイル数", "探索エラー件数", "スキップしたリンク数", _
        "除外したフォルダ数", "終了時刻", "中断・エラー内容", "対象ファイル数", "処理済みファイル数", "進捗率")
    values = Array(status, root, extensionPattern, filePattern, folderPattern, excludedPattern, CStr(m_Written), _
        CStr(m_Examined), CStr(m_ErrorCount), CStr(m_SkippedLinks), CStr(m_ExcludedFolders), Format$(Now, "yyyy-mm-dd hh:nn:ss"), m_LastError, CStr(m_Total), CStr(m_Completed), CStr(Int(InventoryCommon.LastProgressPercent())) & "%")
    For i = 0 To UBound(labels)
        m_Info.Cells(i + 2, 1).Value2 = labels(i)
        m_Info.Cells(i + 2, 2).Value2 = values(i)
    Next i
End Sub

Private Sub FormatReport()
    Dim sheet As Worksheet, width As Long, lastRow As Long, tableRange As Range
    For Each sheet In m_Report.Worksheets
        If sheet.Name = "ファイル一覧" Then
            width = COLUMN_COUNT: lastRow = m_Written + 1
        ElseIf sheet.Name = "探索エラー" Then
            width = 2: lastRow = m_ErrorCount + 1
            If lastRow > sheet.Rows.Count Then lastRow = sheet.Rows.Count
        Else
            width = 2: lastRow = m_Info.Cells(m_Info.Rows.Count, 1).End(xlUp).Row
        End If
        Set tableRange = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, width))
        With sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, width))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
        End With
        tableRange.Columns.AutoFit
        tableRange.AutoFilter
    Next sheet
    m_Files.Columns("B:F").ColumnWidth = 32
    m_Errors.Columns("A:B").ColumnWidth = 55
    m_Info.Columns("B").ColumnWidth = 65
    If m_Written > 1 Then
        m_Files.Range("A1:H" & CStr(m_Written + 1)).Sort _
            Key1:=m_Files.Range("E2"), Order1:=xlAscending, Header:=xlYes
        m_Files.Range("A2:A" & CStr(m_Written + 1)).Value2 = m_Files.Evaluate("ROW(1:" & CStr(m_Written) & ")")
    End If
End Sub
