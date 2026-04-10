Option Explicit

Sub Auto_Close()
    ThisWorkbook.Saved = True   '変更を保存しない
End Sub

'Nullをゼロに変換
Public Function fncNZ(ByVal i_Value As Variant) As Long
    Dim l_Value     As Long
    
    If Not IsNull(i_Value) Then
        l_Value = CLng(i_Value)
    End If

    fncNZ = l_Value
End Function

Public Sub subBeforeEdit()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub subAfterEdit()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Public Function fncNLSP(ByVal str As Variant) As String
    fncNLSP = ""
    If IsNull(str) Then Exit Function
    fncNLSP = CStr(str)
End Function
