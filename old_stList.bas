Option Explicit

Private Const cnTitle As Long = 3

Private Sub cmdEnd_Click()
    If Workbooks.Count = 1 Then Application.Quit
    ThisWorkbook.Close False
End Sub

Private Sub cmdSelect_Click()
    Call subOpenSelect
    If Not P_Regist Then Exit Sub
    Call subMain
End Sub

Private Sub subOpenSelect()
    Dim obj As New frmSelect
    obj.Show
    Set obj = Nothing
End Sub
