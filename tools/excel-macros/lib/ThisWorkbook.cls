Option Explicit

Private Sub Workbook_Open()
    FileInventory.StartFileInventory
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Cancel = True
    Me.Saved = True
    Application.StatusBar = "実行用テンプレートは保存しません。結果ブックを保存してください。"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Me.Saved = True
End Sub
