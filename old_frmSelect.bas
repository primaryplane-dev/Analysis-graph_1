Option Explicit

Private Sub cmbKubun2_Change()

End Sub

Private Sub cmbKubun3_Change()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub UserForm_Initialize()
    P_Regist = False
    optEn.Enabled = True
    optBo.Enabled = True
    cmbKubun3.Enabled = True
    If P_Kengen = 0 Then Call subGetIniFile
    
    Call subMakeCombo
    If P_DateF = 0 Then P_DateF = Date - Day(Date) + 1
    If P_DateT = 0 Then P_DateT = Date
    If P_SelKAI = 0 Then P_SelKAI = 999
    If P_SelKJN = 0 Then P_SelKJN = 999
    
    chkDate.Value = False
    Call subDateEnable
    If P_bDate Then chkDate.Value = True
    cmbYY1.Text = Format(P_DateF, "yyyy")
    cmbMM1.Text = Format(P_DateF, "m")
    cmbDD1.Text = Format(P_DateF, "d")
    cmbYY2.Text = Format(P_DateT, "yyyy")
    cmbMM2.Text = Format(P_DateT, "m")
    cmbDD2.Text = Format(P_DateT, "d")
    
    If Not P_Kubun = 0 Then cmbKubun.Value = P_Kubun
    If Not P_Kubun2 = 0 Then cmbKubun2.Value = P_Kubun2
    
    If Not P_SelKJN = 0 Then cmbKAI.Value = P_SelKAI
    Call subMakeKojoCombo
    If Not P_SelKJN = 0 Then cmbKJN.Value = P_SelKJN
    
    optEn.Value = True
    If P_Type = 2 Then optBo.Value = True
    
    cmdRegist.SetFocus
End Sub

Private Sub chkDate_Click()
    Call subDateEnable
End Sub

Private Sub subDateEnable()
    cmbYY1.Enabled = chkDate.Value
    cmbMM1.Enabled = chkDate.Value
    cmbDD1.Enabled = chkDate.Value
    cmbYY2.Enabled = chkDate.Value
    cmbMM2.Enabled = chkDate.Value
    cmbDD2.Enabled = chkDate.Value
End Sub

Private Sub subMakeCombo()
    Dim vKey        As Variant
    Dim i           As Long
    
    For i = -3 To 0
        cmbYY1.AddItem Year(Date) + i
        cmbYY2.AddItem Year(Date) + i
    Next
    For i = 1 To 12
        cmbMM1.AddItem i
        cmbMM2.AddItem i
    Next
    For i = 1 To 31
        cmbDD1.AddItem i
        cmbDD2.AddItem i
    Next
    
    cmbKubun.Clear
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "1"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "男女比"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "2"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "勤続年数比"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "3"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "男女毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "4"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "勤続年毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "5"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "役職毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "6"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "会社毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "7"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "工場毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "8"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "所属毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "91"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "月毎の"
    cmbKubun.AddItem
    cmbKubun.List(cmbKubun.ListCount - 1, 0) = "92"
    cmbKubun.List(cmbKubun.ListCount - 1, 1) = "年毎の"
    
    
    cmbKubun2.Clear
    cmbKubun2.AddItem
    cmbKubun2.List(cmbKubun2.ListCount - 1, 0) = "1"
    cmbKubun2.List(cmbKubun2.ListCount - 1, 1) = "退職を考えたきっかけ"
    cmbKubun2.AddItem
    cmbKubun2.List(cmbKubun2.ListCount - 1, 0) = "2"
    cmbKubun2.List(cmbKubun2.ListCount - 1, 1) = "次のステップに求めていること"
    
    Call subSetKaisya
End Sub

Private Sub subSetKaisya()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT Distinct"
    strSQL = strSQL & "       RTKAIC"
    strSQL = strSQL & "      ,TRIM(RTKAIM) AS RTKAIM"
    strSQL = strSQL & "      ,RIGHT('00' || TRIM(RTKAIC), 2)"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'メモリが飛んだのに書かれたレコード
    strSQL = strSQL & " ORDER BY 3"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    cmbKAI.Clear
    cmbKAI.AddItem
    cmbKAI.List(cmbKAI.ListCount - 1, 0) = "999"
    cmbKAI.List(cmbKAI.ListCount - 1, 1) = "指定なし"
    Do While Not RS.EOF
        cmbKAI.AddItem
        cmbKAI.List(cmbKAI.ListCount - 1, 0) = Trim(RS("RTKAIC"))
        cmbKAI.List(cmbKAI.ListCount - 1, 1) = Trim(RS("RTKAIM"))
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Sub

Private Sub cmbKubun_Change()
    If cmbKubun.ListIndex = -1 Then Exit Sub
    cmbKubun2.Enabled = True
    If cmbKubun.Value = "1" Or cmbKubun.Value = "2" Or cmbKubun.Value = "91" Or cmbKubun.Value = "92" Then cmbKubun2.ListIndex = -1: cmbKubun2.Enabled = False
    If cmbKubun.Value = "1" Or cmbKubun.Value = "2" Or cmbKubun.Value = "91" Or cmbKubun.Value = "92" Then cmbKubun3.ListIndex = -1: cmbKubun3.Enabled = False
    If cmbKubun.Value = "91" Or cmbKubun.Value = "92" Then optEn.Enabled = False
    
    If cmbKubun.Value = "3" Or cmbKubun.Value = "4" Or cmbKubun.Value = "5" Or cmbKubun.Value = "6" Or cmbKubun.Value = "7" Or cmbKubun.Value = "8" Then
       cmbKubun3.Clear
       cmbKubun3.AddItem
       cmbKubun3.List(cmbKubun3.ListCount - 1, 0) = "1"
       cmbKubun3.List(cmbKubun3.ListCount - 1, 1) = "3年推移"
       cmbKubun3.AddItem
       cmbKubun3.List(cmbKubun3.ListCount - 1, 0) = "2"
       cmbKubun3.List(cmbKubun3.ListCount - 1, 1) = "1年間の月別推移"
       cmbKubun3.Enabled = True
    End If
    
End Sub

Private Sub cmbKAI_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cmbKJN.Clear
    If cmbKAI.ListIndex = -1 Then Exit Sub
    If cmbKAI.Value = "999" Then Exit Sub
    Call subMakeKojoCombo
End Sub

Private Sub subMakeKojoCombo()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT Distinct"
    strSQL = strSQL & "       RTKJNO"
    strSQL = strSQL & "      ,RTKJNM"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'メモリが飛んだのに書かれたレコード
    strSQL = strSQL & "   AND RTKAIC = " & cmbKAI.Value
    strSQL = strSQL & " ORDER BY RTKJNO"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    cmbKJN.Clear
    cmbKJN.AddItem
    cmbKJN.List(cmbKJN.ListCount - 1, 0) = "999"
    cmbKJN.List(cmbKJN.ListCount - 1, 1) = "指定なし"
    Do While Not RS.EOF
        cmbKJN.AddItem
        cmbKJN.List(cmbKJN.ListCount - 1, 0) = Trim(RS("RTKJNO"))
        cmbKJN.List(cmbKJN.ListCount - 1, 1) = Trim(RS("RTKJNM"))
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Sub

Private Sub cmdRegist_Click()
    If cmbKubun.ListIndex = -1 Then Exit Sub
    'If cmbKAI.ListIndex = -1 Then Exit Sub
    'If cmbKJN.ListIndex = -1 Then cmbKJN.ListIndex = 0
    If Not cmbKubun.Value = "1" And Not cmbKubun.Value = "2" And Not cmbKubun.Value = "91" And Not cmbKubun.Value = "92" Then
        If cmbKubun2.ListIndex = -1 Then
            cmbKubun2.SetFocus
            MsgBox "条件選択エラー!!", vbCritical
            Exit Sub
        End If
    End If
    
    If chkDate.Value Then
        If Not IsDate(cmbYY1.Text & "/" & cmbMM1.Text & "/" & cmbDD1.Text) Then
            MsgBox "日付エラー!!", vbCritical
            Exit Sub
        End If
        If Not IsDate(cmbYY2.Text & "/" & cmbMM2.Text & "/" & cmbDD2.Text) Then
            MsgBox "日付エラー!!", vbCritical
            Exit Sub
        End If
    End If
    
    P_Regist = True
    
    P_bDate = chkDate.Value
    P_DateF = CDate(cmbYY1.Text & "/" & cmbMM1.Text & "/" & cmbDD1.Text)
    P_DateT = CDate(cmbYY2.Text & "/" & cmbMM2.Text & "/" & cmbDD2.Text)
    
    P_Kubun = cmbKubun.Value
    If Not cmbKubun.Value = "1" And Not cmbKubun.Value = "2" And Not cmbKubun.Value = "91" And Not cmbKubun.Value = "92" Then
        P_Kubun2 = cmbKubun2.Value
    Else
        P_Kubun2 = 0
    End If
    If cmbKubun3.ListIndex = -1 Then
       P_Kubun3 = 0
    Else
       P_Kubun3 = cmbKubun3.Value
    End If
    
    P_SelKAI = cmbKAI.Value
    P_SelKJN = cmbKJN.Value
    
    If optEn.Value Then P_Type = 1
    If optBo.Value Then P_Type = 2
    
    Unload Me
End Sub
