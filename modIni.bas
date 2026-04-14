Sub ExportYakushokuIniFromMPDP01()
    Dim cn As Object, rs As Object
    Dim strSQL As String
    Dim iniPath As String
    Dim fso As Object, ts As Object

    iniPath = ThisWorkbook.Path & "\役職.ini"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(iniPath, True, False)

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "P_ConnectString" 

    strSQL = "SELECT PDSKUC, PDTKYU FROM LIBMMF.MPDP01 WHERE PDSKUC IS NOT NULL"

    Set rs = cn.Execute(strSQL)

    Do Until rs.EOF
        Dim shokui As String, grade As String, yakushoku As String
        shokui = Trim(rs.Fields("PDSKUC").Value)
        grade = ""
        If Not IsNull(rs.Fields("PDTKYU").Value) Then grade = Trim(rs.Fields("PDTKYU").Value)
        yakushoku = ""

        Select Case shokui
            Case "10"
                yakushoku = "01.役員"
            Case "20"
                yakushoku = "02.役員"
            Case "30"
                If grade = "31" Then
                    yakushoku = "03.課長"
                ElseIf grade = "32" Then
                    yakushoku = "04.統括・次長"
                ElseIf grade = "33" Then
                    yakushoku = "05.部長"
                Else
                    yakushoku = "06.課長等(等級不明)"
                End If
            Case "40"
                If IsNumeric(grade) Then
                    If Val(grade) <= 21 Then
                        yakushoku = "一般"
                    ElseIf Val(grade) = 22 Then
                        yakushoku = "主事補"
                    ElseIf Val(grade) = 23 Then
                        yakushoku = "係長"
                    Else
                        yakushoku = "一般等(等級不明)"
                    End If
                Else
                    yakushoku = "一般等(等級不明)"
                End If
            Case "50"
                yakushoku = "11.研修生"
            Case "60"
                yakushoku = "12.シニア"
            Case "70"
                yakushoku = "13.派遣"
            Case "80"
                yakushoku = "14.パートナー（日給）"
            Case "90"
                yakushoku = "15.パートナー（時給）"
        End Select

        ' ini出力
        If yakushoku <> "" Then
            If shokui = "30" And (grade = "31" Or grade = "32" Or grade = "33") Then
                ts.WriteLine shokui & "," & grade & "," & yakushoku
            ElseIf shokui = "40" And IsNumeric(grade) And Val(grade) <= 21 Then
                ts.WriteLine shokui & ",<=21," & yakushoku
            ElseIf shokui = "40" And grade = "22" Then
                ts.WriteLine shokui & "," & grade & "," & yakushoku
            ElseIf shokui = "40" And grade = "23" Then
                ts.WriteLine shokui & "," & grade & "," & yakushoku
            ElseIf shokui = "40" And (grade = "" Or Not IsNumeric(grade)) Then
                ts.WriteLine shokui & ",*," & yakushoku
            ElseIf shokui = "30" And (grade = "" Or Not (grade = "31" Or grade = "32" Or grade = "33")) Then
                ts.WriteLine shokui & "," & yakushoku
            ElseIf shokui = "10" Or shokui = "20" Or shokui = "50" Or shokui = "60" Or shokui = "70" Or shokui = "80" Or shokui = "90" Then
                ts.WriteLine shokui & "," & yakushoku
            End If
        End If

        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
    ts.Close: Set ts = Nothing: Set fso = Nothing

    MsgBox "役職.iniを出力しました: " & iniPath, vbInformation
End Sub