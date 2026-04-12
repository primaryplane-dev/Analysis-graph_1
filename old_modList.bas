Option Explicit

Private Const cnTitle   As Long = 3
Private arrTitle()      As String
Private arrRow()        As String

Public Sub subMain()
    If P_Kubun3 = 0 Then
        Call subBeforeEdit
        Call subArrTitle
        Call subEditList
        Call subMakeChart
        Call subAfterEdit
    Else
        Call subBeforeEdit
        Call subArrTitle_2
        Call subEditList_2
        Call subMakeChart
        Call subAfterEdit
    End If
End Sub

Private Sub subInitialize(ByRef ST As Worksheet)
    ST.Rows(cnTitle & ":" & Application.Rows.Count).Delete
    ST.Select
End Sub

Private Sub subEditList()
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim i           As Long
    Dim j           As Long
    Dim num         As Long
    
    '雛型シート→編集シート
    Call subInitialize(ST)
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "      RTSEXC"
    strSQL = strSQL & "     ,RTNDAT"
    strSQL = strSQL & "     ,RTTDAT"
    strSQL = strSQL & "     ,RTKNIN"
    strSQL = strSQL & "     ,RTKNIJ"
    strSQL = strSQL & "     ,RTKNID"
    strSQL = strSQL & "     ,RTKNIB"
    strSQL = strSQL & "     ,RTKGYN"
    strSQL = strSQL & "     ,RTKGYR"
    strSQL = strSQL & "     ,RTKZAN"
    strSQL = strSQL & "     ,RTKKYU"
    strSQL = strSQL & "     ,RTKYAS"
    strSQL = strSQL & "     ,RTKYAN"
    strSQL = strSQL & "     ,RTKKIN"
    strSQL = strSQL & "     ,RTKSYO"
    strSQL = strSQL & "     ,RTKHYO"
    strSQL = strSQL & "     ,RTKROU"
    strSQL = strSQL & "     ,RTKKIT"
    strSQL = strSQL & "     ,RTKKAI"
    strSQL = strSQL & "     ,RTKKEK"
    strSQL = strSQL & "     ,RTKSYU"
    strSQL = strSQL & "     ,RTKIKU"
    strSQL = strSQL & "     ,RTKKAG"
    strSQL = strSQL & "     ,RTKKEN"
    strSQL = strSQL & "     ,RTKHAI"
    strSQL = strSQL & "     ,RTKRYU"
    strSQL = strSQL & "     ,RTKSON"
    strSQL = strSQL & "     ,RTNNIN"
    strSQL = strSQL & "     ,RTNGYN"
    strSQL = strSQL & "     ,RTNSEI"
    strSQL = strSQL & "     ,RTNGYR"
    strSQL = strSQL & "     ,RTNZAN"
    strSQL = strSQL & "     ,RTNKYU"
    strSQL = strSQL & "     ,RTNYAS"
    strSQL = strSQL & "     ,RTNKIN"
    strSQL = strSQL & "     ,RTNROU"
    strSQL = strSQL & "     ,RTNJUU"
    strSQL = strSQL & "     ,RTNSYO"
    strSQL = strSQL & "     ,RTNHYO"
    strSQL = strSQL & "     ,RTNKAI"
    strSQL = strSQL & "     ,RTNSON"
    strSQL = strSQL & "     ,RTSKUC"
    strSQL = strSQL & "     ,RTTKYU"
    strSQL = strSQL & "     ,RTKAIM"
    strSQL = strSQL & "     ,RTKJNM"
    strSQL = strSQL & "     ,RTSZBM"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外
    If P_bDate Then strSQL = strSQL & "   AND RTIVD1 BETWEEN " & Format(P_DateF, "yyyymmdd") & " AND " & Format(P_DateT, "yyyymmdd")
    If Not P_SelKAI = 999 Then strSQL = strSQL & "   AND RTKAIC = '" & P_SelKAI & "'"
    If Not P_SelKJN = 999 Then strSQL = strSQL & "   AND RTKJNO = '" & P_SelKJN & "'"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    For i = 0 To UBound(arrTitle)
        ST.Cells(cnTitle + 0, i + 3) = arrTitle(i)
        ST.Cells(cnTitle + 0, i + 3).Font.Size = 12
    Next
    For i = 0 To UBound(arrRow)
        ST.Cells(cnTitle + i + 1, 2) = arrRow(i)
        ST.Cells(cnTitle + i + 1, 2).Font.Size = 12
        num = cnTitle + i + 1
        If P_Kubun = 91 Or P_Kubun = 92 Then
           If num Mod 2 = 0 Then
             ST.Cells(cnTitle + i + 1, 2).Interior.Color = RGB(220, 220, 220)
             ST.Cells(cnTitle + i + 1, 3).Interior.Color = RGB(220, 220, 220)
             ST.Cells(cnTitle + i + 1, 4).Interior.Color = RGB(220, 220, 220)
            End If
        End If
        For j = 0 To UBound(arrTitle)
            ST.Cells(cnTitle + i + 1, j + 3) = 0
            ST.Cells(cnTitle + i + 1, j + 3).Font.Size = 12
        Next
    Next
    
    Do While Not RS.EOF
        lRow = 0
        lCol = 0
        If P_Kubun = 1 Then
            lRow = cnTitle + 1
            If RS("RTSEXC") = "1" Then
                lCol = 3
            ElseIf RS("RTSEXC") = "2" Then
                lCol = 4
            End If
        ElseIf P_Kubun = 2 Then
            lRow = cnTitle + 1
            If CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -10, Date) Then
                lCol = 7
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -5, Date) Then
                lCol = 6
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -3, Date) Then
                lCol = 5
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -1, Date) Then
                lCol = 4
            Else
                lCol = 3
            End If
        ElseIf P_Kubun = 3 Then
            If RS("RTSEXC") = "1" Then
                lRow = cnTitle + 1
            ElseIf RS("RTSEXC") = "2" Then
                lRow = cnTitle + 2
            End If
        ElseIf P_Kubun = 4 Then
            If CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -10, Date) Then
                lRow = cnTitle + 5
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -5, Date) Then
                lRow = cnTitle + 4
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -3, Date) Then
                lRow = cnTitle + 3
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -1, Date) Then
                lRow = cnTitle + 2
            Else
                lRow = cnTitle + 1
            End If
        ElseIf P_Kubun = 5 Then
            If RS("RTSKUC") = "10" Then lRow = cnTitle + 1
            If RS("RTSKUC") = "20" Then lRow = cnTitle + 1
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "31" Then lRow = cnTitle + 2
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "32" Then lRow = cnTitle + 3
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "33" Then lRow = cnTitle + 4
            If RS("RTSKUC") = "40" And RS("RTTKYU") <= "21" Then lRow = cnTitle + 5
            If RS("RTSKUC") = "40" And RS("RTTKYU") = "22" Then lRow = cnTitle + 6
            If RS("RTSKUC") = "40" And RS("RTTKYU") = "23" Then lRow = cnTitle + 7
            If RS("RTSKUC") = "50" Then lRow = cnTitle + 8
            If RS("RTSKUC") = "60" Then lRow = cnTitle + 9
            If RS("RTSKUC") = "70" Then lRow = cnTitle + 10
            If RS("RTSKUC") = "80" Then lRow = cnTitle + 11
            If RS("RTSKUC") = "90" Then lRow = cnTitle + 11
        ElseIf P_Kubun = 6 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTKAIM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        ElseIf P_Kubun = 7 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTKJNM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        ElseIf P_Kubun = 8 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTSZBM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        ElseIf P_Kubun = 91 Then
            If RS("RTSEXC") = "1" Then
               lCol = 3
            ElseIf RS("RTSEXC") = "2" Then
               lCol = 4
            End If
            For i = 0 To 47
              If Left(Trim(RS("RTTDAT")), 4) & "/" & Mid(Trim(RS("RTTDAT")), 5, 2) = arrRow(i) Then
                lRow = cnTitle + 1 + i
              End If
            Next i
        ElseIf P_Kubun = 92 Then
            If RS("RTSEXC") = "1" Then
               lCol = 3
            ElseIf RS("RTSEXC") = "2" Then
               lCol = 4
            End If
            For i = 0 To 3
                If Left(Trim(RS("RTTDAT")), 4) = arrRow(i) Then
                   lRow = cnTitle + 1 + i
                End If
            Next i
        End If
        If P_Kubun2 = 1 Then
            If RS("RTKNIN") = "1" Then lCol = 3
            If RS("RTKGYN") = "1" Then lCol = 4
            If RS("RTKGYR") = "1" Then lCol = 5
            If RS("RTKZAN") = "1" Then lCol = 6
            If RS("RTKKYU") = "1" Then lCol = 7
            If RS("RTKYAS") = "1" Then lCol = 8
            If RS("RTKYAN") = "1" Then lCol = 9
            If RS("RTKKIN") = "1" Then lCol = 10
            If RS("RTKSYO") = "1" Then lCol = 11
            If RS("RTKHYO") = "1" Then lCol = 12
            If RS("RTKROU") = "1" Then lCol = 13
            If RS("RTKKIT") = "1" Then lCol = 14
            If RS("RTKKAI") = "1" Then lCol = 15
            If RS("RTKKEK") = "1" Then lCol = 16
            If RS("RTKSYU") = "1" Then lCol = 17
            If RS("RTKIKU") = "1" Then lCol = 18
            If RS("RTKKAG") = "1" Then lCol = 19
            If RS("RTKKEN") = "1" Then lCol = 20
            If RS("RTKHAI") = "1" Then lCol = 21
            If RS("RTKRYU") = "1" Then lCol = 22
            If RS("RTKSON") = "1" Then lCol = 23
        ElseIf P_Kubun2 = 2 Then
            If RS("RTNNIN") = "1" Then lCol = 3
            If RS("RTNGYN") = "1" Then lCol = 4
            If RS("RTNSEI") = "1" Then lCol = 5
            If RS("RTNGYR") = "1" Then lCol = 6
            If RS("RTNZAN") = "1" Then lCol = 7
            If RS("RTNKYU") = "1" Then lCol = 8
            If RS("RTNYAS") = "1" Then lCol = 9
            If RS("RTNKIN") = "1" Then lCol = 10
            If RS("RTNJUU") = "1" Then lCol = 11
            If RS("RTNSYO") = "1" Then lCol = 12
            If RS("RTNHYO") = "1" Then lCol = 13
            If RS("RTNKAI") = "1" Then lCol = 14
            If RS("RTNSON") = "1" Then lCol = 15
        End If
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, lCol) = ST.Cells(lRow, lCol) + 1
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, 30) = ST.Cells(lRow, 30) + 1
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, lCol).Font.Size = 12
        RS.MoveNext
    Loop
    ST.Range(ST.Cells(cnTitle, 2), ST.Cells(cnTitle + 1, UBound(arrTitle) + 3)).Borders.LineStyle = xlContinuous
    If Not P_Kubun = 1 And Not P_Kubun = 2 Then ST.Range(ST.Cells(cnTitle, 2), ST.Cells(cnTitle + UBound(arrRow) + 1, UBound(arrTitle) + 3)).Borders.LineStyle = xlContinuous
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    ST.Cells(3, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
End Sub

Private Sub subEditList_2()
    Dim ST          As Worksheet: Set ST = stList
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim lRow        As Long
    Dim lCol        As Long
    Dim i           As Long
    Dim j           As Long
    
    '雛型シート→編集シート
    Call subInitialize(ST)
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "      RTSEXC"
    strSQL = strSQL & "     ,RTNDAT"
    strSQL = strSQL & "     ,RTTDAT"
    strSQL = strSQL & "     ,RTKNIN"
    strSQL = strSQL & "     ,RTKNIJ"
    strSQL = strSQL & "     ,RTKNID"
    strSQL = strSQL & "     ,RTKNIB"
    strSQL = strSQL & "     ,RTKGYN"
    strSQL = strSQL & "     ,RTKGYR"
    strSQL = strSQL & "     ,RTKZAN"
    strSQL = strSQL & "     ,RTKKYU"
    strSQL = strSQL & "     ,RTKYAS"
    strSQL = strSQL & "     ,RTKYAN"
    strSQL = strSQL & "     ,RTKKIN"
    strSQL = strSQL & "     ,RTKSYO"
    strSQL = strSQL & "     ,RTKHYO"
    strSQL = strSQL & "     ,RTKROU"
    strSQL = strSQL & "     ,RTKKIT"
    strSQL = strSQL & "     ,RTKKAI"
    strSQL = strSQL & "     ,RTKKEK"
    strSQL = strSQL & "     ,RTKSYU"
    strSQL = strSQL & "     ,RTKIKU"
    strSQL = strSQL & "     ,RTKKAG"
    strSQL = strSQL & "     ,RTKKEN"
    strSQL = strSQL & "     ,RTKHAI"
    strSQL = strSQL & "     ,RTKRYU"
    strSQL = strSQL & "     ,RTKSON"
    strSQL = strSQL & "     ,RTNNIN"
    strSQL = strSQL & "     ,RTNGYN"
    strSQL = strSQL & "     ,RTNSEI"
    strSQL = strSQL & "     ,RTNGYR"
    strSQL = strSQL & "     ,RTNZAN"
    strSQL = strSQL & "     ,RTNKYU"
    strSQL = strSQL & "     ,RTNYAS"
    strSQL = strSQL & "     ,RTNKIN"
    strSQL = strSQL & "     ,RTNROU"
    strSQL = strSQL & "     ,RTNJUU"
    strSQL = strSQL & "     ,RTNSYO"
    strSQL = strSQL & "     ,RTNHYO"
    strSQL = strSQL & "     ,RTNKAI"
    strSQL = strSQL & "     ,RTNSON"
    
    strSQL = strSQL & "     ,RTSKUC"
    strSQL = strSQL & "     ,RTTKYU"
    strSQL = strSQL & "     ,RTKAIM"
    strSQL = strSQL & "     ,RTKJNM"
    strSQL = strSQL & "     ,RTSZBM"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外
    strSQL = strSQL & "   AND RTIVD1 BETWEEN " & "20220101" & " AND " & "20251231"
    If Not P_SelKAI = 999 Then strSQL = strSQL & "   AND RTKAIC = '" & P_SelKAI & "'"
    If Not P_SelKJN = 999 Then strSQL = strSQL & "   AND RTKJNO = '" & P_SelKJN & "'"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    For i = 0 To UBound(arrTitle)
        ST.Cells(cnTitle + 0, i + 3) = arrTitle(i)
        ST.Cells(cnTitle + 0, i + 3).Font.Size = 12
    Next
    For i = 0 To UBound(arrRow)
        ST.Cells(cnTitle + i + 1, 2) = arrRow(i)
        ST.Cells(cnTitle + i + 1, 2).Font.Size = 12
        For j = 0 To UBound(arrTitle)
            ST.Cells(cnTitle + i + 1, j + 3) = 0
            ST.Cells(cnTitle + i + 1, j + 3).Font.Size = 12
        Next
    Next
    If P_Kubun = 3 And P_Kubun3 = 1 Then
       'Rows(cnTitle).Copy Destination:=Rows(cnTitle + 5)
    End If
    
    Do While Not RS.EOF
        lRow = 0
        lCol = 0
        
        If P_Kubun = 3 Then
            If RS("RTSEXC") = "1" Then
                If Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 3) Then
                   lRow = cnTitle + 1
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 2) Then
                   lRow = cnTitle + 2
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 1) Then
                   lRow = cnTitle + 3
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date)) Then
                   lRow = cnTitle + 4
                End If
            ElseIf RS("RTSEXC") = "2" Then
                If Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 3) Then
                   lRow = cnTitle + 5
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 2) Then
                   lRow = cnTitle + 6
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date) - 1) Then
                   lRow = cnTitle + 7
                ElseIf Left(RS("RTTDAT"), 4) = Trim(Year(Date)) Then
                   lRow = cnTitle + 8
                End If
            End If
        ElseIf P_Kubun = 4 Then
            If CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -10, Date) Then
                lRow = cnTitle + 5
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -5, Date) Then
                lRow = cnTitle + 4
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -3, Date) Then
                lRow = cnTitle + 3
            ElseIf CDate(Format(RS("RTNDAT"), "0000/00/00")) <= DateAdd("yyyy", -1, Date) Then
                lRow = cnTitle + 2
            Else
                lRow = cnTitle + 1
            End If
        ElseIf P_Kubun = 5 Then
            If RS("RTSKUC") = "10" Then lRow = cnTitle + 1
            If RS("RTSKUC") = "20" Then lRow = cnTitle + 1
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "31" Then lRow = cnTitle + 2
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "32" Then lRow = cnTitle + 3
            If RS("RTSKUC") = "30" And RS("RTTKYU") = "33" Then lRow = cnTitle + 4
            If RS("RTSKUC") = "40" And RS("RTTKYU") <= "21" Then lRow = cnTitle + 5
            If RS("RTSKUC") = "40" And RS("RTTKYU") = "22" Then lRow = cnTitle + 6
            If RS("RTSKUC") = "40" And RS("RTTKYU") = "23" Then lRow = cnTitle + 7
            If RS("RTSKUC") = "50" Then lRow = cnTitle + 8
            If RS("RTSKUC") = "60" Then lRow = cnTitle + 9
            If RS("RTSKUC") = "70" Then lRow = cnTitle + 10
            If RS("RTSKUC") = "80" Then lRow = cnTitle + 11
            If RS("RTSKUC") = "90" Then lRow = cnTitle + 11
        ElseIf P_Kubun = 6 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTKAIM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        ElseIf P_Kubun = 7 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTKJNM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        ElseIf P_Kubun = 8 Then
            For i = 0 To UBound(arrRow)
                If Trim(RS("RTSZBM")) = arrRow(i) Then
                    lRow = cnTitle + i + 1
                End If
            Next
        
        End If
        If P_Kubun2 = 1 Then
            If RS("RTKNIN") = "1" Then lCol = 3
            If RS("RTKGYN") = "1" Then lCol = 4
            If RS("RTKGYR") = "1" Then lCol = 5
            If RS("RTKZAN") = "1" Then lCol = 6
            If RS("RTKKYU") = "1" Then lCol = 7
            If RS("RTKYAS") = "1" Then lCol = 8
            If RS("RTKYAN") = "1" Then lCol = 9
            If RS("RTKKIN") = "1" Then lCol = 10
            If RS("RTKSYO") = "1" Then lCol = 11
            If RS("RTKHYO") = "1" Then lCol = 12
            If RS("RTKROU") = "1" Then lCol = 13
            If RS("RTKKIT") = "1" Then lCol = 14
            If RS("RTKKAI") = "1" Then lCol = 15
            If RS("RTKKEK") = "1" Then lCol = 16
            If RS("RTKSYU") = "1" Then lCol = 17
            If RS("RTKIKU") = "1" Then lCol = 18
            If RS("RTKKAG") = "1" Then lCol = 19
            If RS("RTKKEN") = "1" Then lCol = 20
            If RS("RTKHAI") = "1" Then lCol = 21
            If RS("RTKRYU") = "1" Then lCol = 22
            If RS("RTKSON") = "1" Then lCol = 23
        ElseIf P_Kubun2 = 2 Then
            If RS("RTNNIN") = "1" Then lCol = 3
            If RS("RTNGYN") = "1" Then lCol = 4
            If RS("RTNSEI") = "1" Then lCol = 5
            If RS("RTNGYR") = "1" Then lCol = 6
            If RS("RTNZAN") = "1" Then lCol = 7
            If RS("RTNKYU") = "1" Then lCol = 8
            If RS("RTNYAS") = "1" Then lCol = 9
            If RS("RTNKIN") = "1" Then lCol = 10
            If RS("RTNJUU") = "1" Then lCol = 11
            If RS("RTNSYO") = "1" Then lCol = 12
            If RS("RTNHYO") = "1" Then lCol = 13
            If RS("RTNKAI") = "1" Then lCol = 14
            If RS("RTNSON") = "1" Then lCol = 15
        End If
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, lCol) = ST.Cells(lRow, lCol) + 1
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, 30) = ST.Cells(lRow, 30) + 1
        If Not lRow = 0 And Not lCol = 0 Then ST.Cells(lRow, lCol).Font.Size = 12
        RS.MoveNext
    Loop
    ST.Range(ST.Cells(cnTitle, 2), ST.Cells(cnTitle + 1, UBound(arrTitle) + 3)).Borders.LineStyle = xlContinuous
    If Not P_Kubun = 1 And Not P_Kubun = 2 Then ST.Range(ST.Cells(cnTitle, 2), ST.Cells(cnTitle + UBound(arrRow) + 1, UBound(arrTitle) + 3)).Borders.LineStyle = xlContinuous
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    ST.Cells(3, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Set ST = Nothing
End Sub




Private Sub subMakeChart()
    Dim i       As Long
    Dim sTitle  As String
    Call subChartDelete
    If P_Kubun = 1 Then
        Call subAddChart("男女比", "男女比", Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle + 1, 4)))
    ElseIf P_Kubun = 2 Then
        Call subAddChart("勤続年数比", "勤続年数比", Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle + 1, 7)))
    ElseIf P_Kubun = 91 Then
        Dim Row As Long
        For i = 3 To 51
         If stList.Cells(i, 30) <> "" Then
            Row = i
            Exit For
         End If
        Next
        Call subAddChart_tukinen("月ごとの退職者数", "月ごとの退職者数", Range(stList.Cells(Row, 2), stList.Cells(cnTitle + 48, 4)))
    ElseIf P_Kubun = 92 Then
        Call subAddChart_tukinen("年ごとの退職者数", "年ごとの退職者数", Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle + 4, 4)))
    Else
        If P_Kubun3 = 0 Then
            If P_Kubun2 = 1 Then sTitle = "退職を考えたきっかけ"
            If P_Kubun2 = 2 Then sTitle = "次のステップに求めていること"
            For i = UBound(arrRow) To 0 Step -1
                If Not stList.Cells(cnTitle + i + 1, 30) = "" Then
                    If i = 0 Then
                        Call subAddChart(sTitle & "(" & arrRow(i) & ")", arrRow(i), Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle + 1, UBound(arrTitle) + 3)))
                    Else
                        Call subAddChart(sTitle & "(" & arrRow(i) & ")", arrRow(i), Union(Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle, UBound(arrTitle) + 3)), Range(stList.Cells(cnTitle + i + 1, 2), stList.Cells(cnTitle + i + 1, UBound(arrTitle) + 3))))
                    End If
                End If
            Next
        ElseIf P_Kubun3 = 1 Then
            If P_Kubun = 3 Then
                If P_Kubun2 = 1 Then sTitle = "退職を考えたきっかけ"
                If P_Kubun2 = 2 Then sTitle = "次のステップに求めていること"
                Call subAddChart(sTitle, "3年推移男", Range(stList.Cells(cnTitle, 2), stList.Cells(cnTitle + 4, UBound(arrTitle) + 3)))
                'Call subAddChart(sTitle, "3年推移女", Range(stList.Cells(cnTitle+5, 2),stList.Cells(cnTitle + 4, UBound(arrTitle) + 3))
             End If
        End If
    End If
End Sub

Private Sub subAddChart(ByVal sTitle As String, ByVal sSheetName As String, ByRef dsRange As Range)
    Dim obj     As Chart
    
    Set obj = ThisWorkbook.Charts.Add
    obj.Move After:=stList
    obj.Name = Replace(StrConv(Trim(sSheetName), vbNarrow), "株式会社", "")
    If P_Type = 1 Then
        obj.ChartType = xlPie
        obj.ApplyDataLabels xlDataLabelsShowPercent
    ElseIf P_Type = 2 Then
        obj.ChartType = xlColumnClustered
        obj.ApplyDataLabels xlDataLabelsShowValue
    End If
    obj.HasTitle = True
    obj.ChartTitle.Text = sTitle
    obj.SetSourceData dsRange
    obj.Protect
End Sub

Private Sub subAddChart_tukinen(ByVal sTitle As String, ByVal sSheetName As String, ByRef dsRange As Range)
   Dim obj As Chart
   Set obj = ThisWorkbook.Charts.Add
   obj.Move After:=stList
   obj.Name = sSheetName
   
   obj.ChartType = xlColumnClustered
   obj.ApplyDataLabels xlDataLabelsShowValue
    
   obj.HasTitle = True
   obj.ChartTitle.Text = sTitle
   obj.SetSourceData dsRange
   obj.Protect

End Sub




Private Sub subChartDelete()
    Dim obj As Chart
    Application.DisplayAlerts = False
    For Each obj In ThisWorkbook.Charts
        obj.Delete
    Next
    Application.DisplayAlerts = True
End Sub

Private Sub subArrTitle()
    If P_Kubun = 1 Then
        ReDim arrTitle(1)
        ReDim arrRow(0)
        arrTitle(0) = "男"
        arrTitle(1) = "女"
        arrRow(0) = "人数"
    ElseIf P_Kubun = 2 Then
        ReDim arrTitle(4)
        ReDim arrRow(0)
        arrTitle(0) = "1年未満"
        arrTitle(1) = "1年以上"
        arrTitle(2) = "3年以上"
        arrTitle(3) = "5年以上"
        arrTitle(4) = "10年以上"
        arrRow(0) = "人数"
    ElseIf P_Kubun = 3 Then
        ReDim arrRow(1)
        arrRow(0) = "男"
        arrRow(1) = "女"
    ElseIf P_Kubun = 4 Then
        ReDim arrRow(4)
        arrRow(0) = "1年未満"
        arrRow(1) = "1年以上"
        arrRow(2) = "3年以上"
        arrRow(3) = "5年以上"
        arrRow(4) = "10年以上"
    ElseIf P_Kubun = 5 Then
        ReDim arrRow(10)
        arrRow(0) = "役員"
        arrRow(1) = "課長"
        arrRow(2) = "統括・次長"
        arrRow(3) = "部長"
        arrRow(4) = "一般"
        arrRow(5) = "主事補"
        arrRow(6) = "係長"
        arrRow(7) = "研修生"
        arrRow(8) = "シニア"
        arrRow(9) = "派遣"
        arrRow(10) = "パートナー"
    ElseIf P_Kubun = 6 Then
        Call subSetKaisya
    ElseIf P_Kubun = 7 Then
        Call subSetKojo
    ElseIf P_Kubun = 8 Then
        Call subSetShozoku
    ElseIf P_Kubun = 91 Then
        ReDim arrTitle(1)
        arrTitle(0) = "男"
        arrTitle(1) = "女"
        Call subSettukigoto
    ElseIf P_Kubun = 92 Then
        ReDim arrTitle(1)
        arrTitle(0) = "男"
        arrTitle(1) = "女"
        Call subSetnengoto
    End If
    If P_Kubun2 = 1 Then
        ReDim arrTitle(20)
        arrTitle(0) = "人間関係の悩み"
        arrTitle(1) = "業務内容にやりがいを感じない"
        arrTitle(2) = "業務量の多さ"
        arrTitle(3) = "残業量の多さ"
        arrTitle(4) = "給与"
        arrTitle(5) = "年間休日が少ない"
        arrTitle(6) = "休日が取れない"
        arrTitle(7) = "勤務地"
        arrTitle(8) = "昇格に不満"
        arrTitle(9) = "評価に不満"
        arrTitle(10) = "環境"
        arrTitle(11) = "勤務時間が合わない"
        arrTitle(12) = "会社の将来性に不安"
        arrTitle(13) = "結婚"
        arrTitle(14) = "出産"
        arrTitle(15) = "育児"
        arrTitle(16) = "介護"
        arrTitle(17) = "健康面に不安がある"
        arrTitle(18) = "配偶者転勤"
        arrTitle(19) = "キャリアアップ"
        arrTitle(20) = "その他"
    ElseIf P_Kubun2 = 2 Then
        ReDim arrTitle(12)
        arrTitle(0) = "人間関係"
        arrTitle(1) = "業務内容"
        arrTitle(2) = "やりがい・成長"
        arrTitle(3) = "業務量"
        arrTitle(4) = "残業量"
        arrTitle(5) = "給与"
        arrTitle(6) = "年間休日の日数や希望"
        arrTitle(7) = "勤務地"
        arrTitle(8) = "柔軟な働き方"
        arrTitle(9) = "昇格"
        arrTitle(10) = "評価"
        arrTitle(11) = "会社の将来性"
        arrTitle(12) = "その他"
    End If
End Sub

Private Sub subArrTitle_2()
       If P_Kubun = 3 Then
          ReDim arrRow(7)
          arrRow(0) = Year(Date) - 3 & "/" & "男"
          arrRow(1) = Year(Date) - 2 & "/" & "男"
          arrRow(2) = Year(Date) - 1 & "/" & "男"
          arrRow(3) = Year(Date) & "/" & "男"
          arrRow(4) = Year(Date) - 3 & "/" & "女"
          arrRow(5) = Year(Date) - 2 & "/" & "女"
          arrRow(6) = Year(Date) - 1 & "/" & "女"
          arrRow(7) = Year(Date) & "/" & "女"
        End If
        
        If P_Kubun2 = 1 Then
            ReDim arrTitle(20)
            arrTitle(0) = "人間関係の悩み"
            arrTitle(1) = "業務内容にやりがいを感じない"
            arrTitle(2) = "業務量の多さ"
            arrTitle(3) = "残業量の多さ"
            arrTitle(4) = "給与"
            arrTitle(5) = "年間休日が少ない"
            arrTitle(6) = "休日が取れない"
            arrTitle(7) = "勤務地"
            arrTitle(8) = "昇格に不満"
            arrTitle(9) = "評価に不満"
            arrTitle(10) = "環境"
            arrTitle(11) = "勤務時間が合わない"
            arrTitle(12) = "会社の将来性に不安"
            arrTitle(13) = "結婚"
            arrTitle(14) = "出産"
            arrTitle(15) = "育児"
            arrTitle(16) = "介護"
            arrTitle(17) = "健康面に不安がある"
            arrTitle(18) = "配偶者転勤"
            arrTitle(19) = "キャリアアップ"
            arrTitle(20) = "その他"
       ElseIf P_Kubun2 = 2 Then
            ReDim arrTitle(12)
            arrTitle(0) = "人間関係"
            arrTitle(1) = "業務内容"
            arrTitle(2) = "やりがい・成長"
            arrTitle(3) = "業務量"
            arrTitle(4) = "残業量"
            arrTitle(5) = "給与"
            arrTitle(6) = "年間休日の日数や希望"
            arrTitle(7) = "勤務地"
            arrTitle(8) = "柔軟な働き方"
            arrTitle(9) = "昇格"
            arrTitle(10) = "評価"
            arrTitle(11) = "会社の将来性"
            arrTitle(12) = "その他"
        End If
'    ElseIf P_Kubun = 4 Then
'        ReDim arrRow(4)
'        arrRow(0) = "1年未満"
'        arrRow(1) = "1年以上"
'        arrRow(2) = "3年以上"
'        arrRow(3) = "5年以上"
'        arrRow(4) = "10年以上"
'    ElseIf P_Kubun = 5 Then
'        ReDim arrRow(10)
'        arrRow(0) = "役員"
'        arrRow(1) = "課長"
'        arrRow(2) = "統括・次長"
'        arrRow(3) = "部長"
'        arrRow(4) = "一般"
'        arrRow(5) = "主事補"
'        arrRow(6) = "係長"
'        arrRow(7) = "研修生"
'        arrRow(8) = "シニア"
'        arrRow(9) = "派遣"
'        arrRow(10) = "パートナー"
'    ElseIf P_Kubun = 6 Then
'        Call subSetKaisya
'    ElseIf P_Kubun = 7 Then
'        Call subSetKojo
'    ElseIf P_Kubun = 8 Then
'        Call subSetShozoku
'    ElseIf P_Kubun = 91 Then
'        ReDim arrTitle(1)
'        arrTitle(0) = "男"
'        arrTitle(1) = "女"
'        Call subSettukigoto
'    ElseIf P_Kubun = 92 Then
'        ReDim arrTitle(1)
'        arrTitle(0) = "男"
'        arrTitle(1) = "女"
'        Call subSetnengoto


End Sub

Private Sub subSetKaisya()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim i           As Long
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT Distinct"
    strSQL = strSQL & "       RTKAIC"
    strSQL = strSQL & "      ,RTKAIM"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外
    strSQL = strSQL & " ORDER BY RTKAIC"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    ReDim arrRow(RS.RecordCount - 1)
    i = 0
    Do While Not RS.EOF
        arrRow(i) = Trim(RS("RTKAIM"))
        i = i + 1
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Sub

Private Sub subSetKojo()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim i           As Long
    
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
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外
    strSQL = strSQL & " ORDER BY RTKJNO"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    ReDim arrRow(RS.RecordCount - 1)
    i = 0
    Do While Not RS.EOF
        arrRow(i) = Trim(RS("RTKJNM"))
        i = i + 1
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Sub

Private Sub subSetShozoku()
    Dim CN          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim strSQL      As String
    Dim i           As Long
    
    If P_Kengen = 0 Then Call subGetIniFile
    
    'ＤＢ接続
    CN.CursorLocation = adUseClient
    CN.Open P_ConnectString
    
    strSQL = ""
    strSQL = strSQL & " SELECT Distinct"
    strSQL = strSQL & "       RTKAIC"
    strSQL = strSQL & "      ,RTKJNO"
    strSQL = strSQL & "      ,RTSZBM"
    strSQL = strSQL & " FROM LIBIMF.IRTP01"
    strSQL = strSQL & " WHERE RTDLT <> 'X'"
    strSQL = strSQL & "   AND NOT RTKJNO = 0"       'RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外
    strSQL = strSQL & " ORDER BY RTKAIC, RTKJNO"

    RS.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly

    ReDim arrRow(RS.RecordCount - 1)
    i = 0
    Do While Not RS.EOF
        arrRow(i) = Trim(RS("RTSZBM"))
        i = i + 1
        RS.MoveNext
    Loop
    
    'ＤＢ切断
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
End Sub

Private Sub subSettukigoto()
  
  '月ごとの退職者数
  ReDim arrRow(47)
  arrRow(0) = Year(Date) - 3 & "/" & "01"
  arrRow(1) = Year(Date) - 3 & "/" & "02"
  arrRow(2) = Year(Date) - 3 & "/" & "03"
  arrRow(3) = Year(Date) - 3 & "/" & "04"
  arrRow(4) = Year(Date) - 3 & "/" & "05"
  arrRow(5) = Year(Date) - 3 & "/" & "06"
  arrRow(6) = Year(Date) - 3 & "/" & "07"
  arrRow(7) = Year(Date) - 3 & "/" & "08"
  arrRow(8) = Year(Date) - 3 & "/" & "09"
  arrRow(9) = Year(Date) - 3 & "/" & "10"
  arrRow(10) = Year(Date) - 3 & "/" & "11"
  arrRow(11) = Year(Date) - 3 & "/" & "12"
  arrRow(12) = Year(Date) - 2 & "/" & "01"
  arrRow(13) = Year(Date) - 2 & "/" & "02"
  arrRow(14) = Year(Date) - 2 & "/" & "03"
  arrRow(15) = Year(Date) - 2 & "/" & "04"
  arrRow(16) = Year(Date) - 2 & "/" & "05"
  arrRow(17) = Year(Date) - 2 & "/" & "06"
  arrRow(18) = Year(Date) - 2 & "/" & "07"
  arrRow(19) = Year(Date) - 2 & "/" & "08"
  arrRow(20) = Year(Date) - 2 & "/" & "09"
  arrRow(21) = Year(Date) - 2 & "/" & "10"
  arrRow(22) = Year(Date) - 2 & "/" & "11"
  arrRow(23) = Year(Date) - 2 & "/" & "12"
  arrRow(24) = Year(Date) - 1 & "/" & "01"
  arrRow(25) = Year(Date) - 1 & "/" & "02"
  arrRow(26) = Year(Date) - 1 & "/" & "03"
  arrRow(27) = Year(Date) - 1 & "/" & "04"
  arrRow(28) = Year(Date) - 1 & "/" & "05"
  arrRow(29) = Year(Date) - 1 & "/" & "06"
  arrRow(30) = Year(Date) - 1 & "/" & "07"
  arrRow(31) = Year(Date) - 1 & "/" & "08"
  arrRow(32) = Year(Date) - 1 & "/" & "09"
  arrRow(33) = Year(Date) - 1 & "/" & "10"
  arrRow(34) = Year(Date) - 1 & "/" & "11"
  arrRow(35) = Year(Date) - 1 & "/" & "12"
  arrRow(36) = Year(Date) & "/" & "01"
  arrRow(37) = Year(Date) & "/" & "02"
  arrRow(38) = Year(Date) & "/" & "03"
  arrRow(39) = Year(Date) & "/" & "04"
  arrRow(40) = Year(Date) & "/" & "05"
  arrRow(41) = Year(Date) & "/" & "06"
  arrRow(42) = Year(Date) & "/" & "07"
  arrRow(43) = Year(Date) & "/" & "08"
  arrRow(44) = Year(Date) & "/" & "09"
  arrRow(45) = Year(Date) & "/" & "10"
  arrRow(46) = Year(Date) & "/" & "11"
  arrRow(47) = Year(Date) & "/" & "12"
  
End Sub
Private Sub subSetnengoto()
  '年ごとの退職者数
  ReDim arrRow(3)
  arrRow(0) = Year(Date) - 3
  arrRow(1) = Year(Date) - 2
  arrRow(2) = Year(Date) - 1
  arrRow(3) = Year(Date)
  
End Sub
