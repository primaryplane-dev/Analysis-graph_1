Option Explicit

Public Const P_ConnectString As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=SYSTEM;Password=FJPN2480"
'Public Const P_ConnectString As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=SYSTEM;Password=FJPN2480"

'
Public Function fnTrimCode(ByVal str As String) As String
    Dim Target As String: Target = Trim(str)
    If Target = "" Then Exit Function
    
    ' 数字2桁 + ドット（半角/全角）があれば、その次からを返す
    If Target Like "##.*" Or Target Like "##．*" Then
        fnTrimCode = Mid(Target, 4)
    Else
        fnTrimCode = Target
    End If
End Function
