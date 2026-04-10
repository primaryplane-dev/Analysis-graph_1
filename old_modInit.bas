Option Explicit

'「設定.ini」の内容を取得する。
Public Sub subGetIniFile()

    Dim L_StrIniFile    As String
    Dim L_StrBuffer     As String
    Dim L_StrKey        As String
    Dim L_StrValue      As String

'    L_StrIniFile = Application.ThisWorkbook.Path & "\設定.ini"
'    Open L_StrIniFile For Input As #1
'    Do While Not EOF(1)
'        Line Input #1, L_StrBuffer
'        If Len(Trim(L_StrBuffer)) > 0 Then
'            Call PS_Split(L_StrBuffer, L_StrKey, L_StrValue)
'            Select Case L_StrKey
'            Case "NYNO":        P_NYNO = L_StrValue
'            Case "Kengen":      P_Kengen = L_StrValue
'            End Select
'        End If
'    Loop
'    Close #1
    
    Call subGetIniFile2
    
End Sub
'文字列分割
Public Sub PS_Split(ByVal I_StrString As String, ByRef O_StrKey As String, ByRef O_StrValue As String)

    Dim L_IntStart  As Integer
    
    L_IntStart = InStr(1, I_StrString, "=")
    O_StrKey = Left(I_StrString, L_IntStart - 1)
    O_StrValue = Mid(I_StrString, L_IntStart + 1)
    
End Sub

'「工場.ini」の内容を取得する。
Private Sub subGetIniFile2()
    Dim L_StrIniFile    As String
    Dim L_StrBuffer     As String
    Dim sSp()           As String
    Dim sKey            As String
    Dim sValue          As String
    
    If Not dicKaisyaName Is Nothing Then Set dicKaisyaName = Nothing
    If Not dicKojoName Is Nothing Then Set dicKojoName = Nothing
    
    Set dicKaisyaName = New Dictionary
    Set dicKojoName = New Dictionary
    
'    L_StrIniFile = Application.ThisWorkbook.Path & "\会社.ini"
'    Open L_StrIniFile For Input As #1
'    Do While Not EOF(1)
'        Line Input #1, L_StrBuffer
'        sSp = Split(L_StrBuffer, ",")
'
'        sKey = CLng(sSp(0))
'        sValue = sSp(1)
'        If Not dicKaisyaName.Exists(sKey) Then dicKaisyaName.Add sKey, sValue
'        sKey = Format(CLng(sSp(0)), "00") & Format(CLng(sSp(2)), "000")
'        sValue = sSp(3)
'        If Not dicKojoName.Exists(sKey) Then dicKojoName.Add sKey, sValue
'    Loop
'    Close #1
End Sub
