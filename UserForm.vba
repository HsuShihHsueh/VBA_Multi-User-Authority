Private Sub CommandButton1_Click()

Dim j, i As Integer

max_y = Application.CountA(Sheets(2).Range("A:A"))
max_x = Application.CountA(Sheets(2).Range("1:1"))

If TextBox1.Value = "" Then MsgBox "用戶名不能為空", vbInformation, "注意": Exit Sub
If TextBox2.Value = "" Then MsgBox "密碼名不能為空", vbInformation, "注意": Exit Sub

For i = 2 To max_x
    Name_1 = Sheets(2).Cells(1, i)
    Password_1 = Sheets(2).Cells(2, i)
        If TextBox1.Text = Name_1 And TextBox2.Text = Password_1 Then
            'user & pssword correct
            Unload Me
            MsgBox ("Welcome! " + Name_1)
            'ThisWorkbook.Activate
            For j = max_y To 3 Step -1
            'y from buttom to 3(1,2 is user and password)
                a = Sheets(2).Cells(j, 1)
                If Sheets(2).Cells(j, i) = "y" Then
                    If a = "全部" Then
                        For Each Sht In Sheets
                            Sht.Visible = True
                        Next
                        Sheets(1).Visible = xlVeryHidden
                        Application.Visible = True
                        Exit Sub    'end process
                    Else
                        Sheets(a).Visible = True
                    End If
                End If
            Next j
            Sheets(1).Visible = xlVeryHidden
            Application.Visible = True
            Exit Sub    'if enter any process the exit
        End If
Next i
MsgBox "帳號或密碼錯誤", vbInformation, "error"
End Sub

Private Sub CommandButton2_Click()
Application.DisplayAlerts = False   'make warning wont pop up
Unload Me                                           'release the ram
Application.Quit                                  'close excel
Application.EnableEvents = False    'make process wont got die loop
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then Cancel = True
End Sub
