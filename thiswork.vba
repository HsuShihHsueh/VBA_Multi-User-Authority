Public strName As String    'determine whether the process is "save as ..."

Private Sub Workbook_Open()
    strName = ActiveWorkbook.Name   'save the inital workbook name
    Sheets(1).Visible = True
    For Each Sht In Sheets
        If Not Sht.Name = Sheets(1).Name Then
            Sht.Visible = xlVeryHidden
        End If
    Next
    Application.DisplayAlerts = False
    Application.Visible = False
    Application.EnableCancelKey = xlDisabled
    UserForm1.Caption = "Login in"
    UserForm1.Show
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Sheets(1).Visible = True
    For k = 2 To Application.Sheets.Count Step 1
        Sheets(k).Visible = xlVeryHidden
    Next k
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Not ActiveWorkbook.Name = strName Then
        Sheets(1).Visible = True
        For k = 2 To Application.Sheets.Count Step 1
            Sheets(k).Visible = xlVeryHidden
        Next k
        MsgBox ("Save as： " + ActiveWorkbook.Name)
        strName = ActiveWorkbook.Name   'save the workbook name after save as
        ActiveWorkbook.Save
    End If
End Sub
