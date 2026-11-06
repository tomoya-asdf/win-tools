Option Explicit

Private Const REPARSE_POINT As Long = 1024
Private m_Progress As Double

Public Sub ResetProgress()
    m_Progress = 0
End Sub

Public Function LastProgressPercent() As Double
    LastProgressPercent = m_Progress
End Function

Public Sub ShowProgress(ByVal completed As Long, ByVal total As Long, ByVal label As String)
    If total = 0 Then
        m_Progress = 100
    Else
        m_Progress = 100# * completed / total
    End If
    Application.StatusBar = label & "  " & CStr(Int(m_Progress)) & "% (" & CStr(completed) & "/" & CStr(total) & ")"
    DoEvents
End Sub

Public Function MakeRegex(ByVal pattern As String, ByVal label As String) As Object
    Dim regex As Object, valid As Boolean, message As String
    If Len(pattern) = 0 Then Exit Function
    On Error GoTo Invalid
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    valid = regex.Test(vbNullString)
    Set MakeRegex = regex
    Exit Function
Invalid:
    message = Err.Description
    Err.Raise vbObjectError + 701, "Inventory", label & "の正規表現を確認してください: " & message
End Function

Public Function Matches(ByVal regex As Object, ByVal value As String) As Boolean
    If regex Is Nothing Then Matches = True Else Matches = regex.Test(value)
End Function

Public Function AskPattern(ByVal prompt As String, ByVal defaultValue As String, ByRef result As String) As Boolean
    Dim answer As Variant
    answer = Application.InputBox(prompt, "検索条件（空欄はすべて・大文字小文字を区別しない）", defaultValue, Type:=2)
    If VarType(answer) = vbBoolean Then Exit Function
    result = CStr(answer)
    AskPattern = True
End Function

Public Function CollectFiles(ByVal root As String, ByVal extensionRegex As Object, ByVal fileRegex As Object, _
    ByVal folderRegex As Object, ByVal excludedRegex As Object, ByRef issues As Collection, _
    ByRef examined As Long, ByRef skipped As Long, ByRef excluded As Long) As Collection
    Dim fs As Object, pending As Collection, files As Collection, path As String, relative As String, prefix As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(root) Then Err.Raise vbObjectError + 702, , "探索フォルダが見つかりません。"
    root = fs.GetFolder(root).Path
    prefix = root
    If Right$(prefix, 1) <> "\" Then prefix = prefix & "\"
    Set pending = New Collection
    Set files = New Collection
    Set issues = New Collection
    pending.Add root
    Do While pending.Count > 0
        path = CStr(pending(pending.Count))
        pending.Remove pending.Count
        If StrComp(root, path, vbTextCompare) = 0 Then relative = "." Else relative = Mid$(path, Len(prefix) + 1)
        Application.StatusBar = "対象を収集中（総数を確認しています）: " & path
        If Not excludedRegex Is Nothing Then
            If excludedRegex.Test(relative) Then
                excluded = excluded + 1
                GoTo NextFolder
            End If
        End If
        CollectFolder fs, path, relative, pending, files, issues, extensionRegex, fileRegex, folderRegex, examined, skipped
NextFolder:
        DoEvents
    Loop
    Set CollectFiles = files
End Function

Private Sub CollectFolder(ByVal fs As Object, ByVal path As String, ByVal relative As String, _
    ByVal pending As Collection, ByVal files As Collection, ByVal issues As Collection, _
    ByVal extensionRegex As Object, ByVal fileRegex As Object, ByVal folderRegex As Object, _
    ByRef examined As Long, ByRef skipped As Long)
    Dim folder As Object, item As Object, child As Object, number As Long, message As String
    On Error GoTo Failed
    Set folder = fs.GetFolder(path)
    If (folder.Attributes And REPARSE_POINT) <> 0 Then
        skipped = skipped + 1
        Exit Sub
    End If
    For Each item In folder.Files
        examined = examined + 1
        If (item.Attributes And REPARSE_POINT) <> 0 Then
            skipped = skipped + 1
        ElseIf Matches(extensionRegex, fs.GetExtensionName(item.Name)) And Matches(fileRegex, item.Name) And Matches(folderRegex, relative) Then
            files.Add Array(item.Path, relative)
        End If
        If examined Mod 100 = 0 Then DoEvents
    Next item
    For Each child In folder.SubFolders
        pending.Add child.Path
    Next child
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    If number = 18 Then Err.Raise number, , message
    issues.Add Array(path, message)
End Sub

Public Sub CloseInventoryTool(Optional ByVal report As Object = Nothing)
    If Not report Is Nothing Then report.Activate
    ThisWorkbook.Close SaveChanges:=False
End Sub
