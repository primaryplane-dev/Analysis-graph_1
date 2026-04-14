Option Explicit

Public Sub subMain()
    Application.ScreenUpdating = False
    
    ' 1. INIファイルの読込（会社・役職・所属の辞書作成）
    Call subGetIniFileAll
    
    ' 2. データ抽出（オフラインのときはコメントにする）
'    Call subExtractDataToSheet
    
    ' 3. ラベル付与と理由集約
    Call subAddCalculatedColumns
    
    ' 4. メニューのパターンに合わせて表示更新（ここでヘッダーやグラフを作る）
    Call subUpdatePivotByMenu(P_Pattern)
    
    ' 画面更新を戻す
    Application.ScreenUpdating = True
    
End Sub

' INIファイル（会社・役職・所属）の読み込み処理
Public Sub subGetIniFileAll()
    ' 辞書の初期化
    Set dicKaisyaName = CreateObject("Scripting.Dictionary")
    Set dicKojoName = CreateObject("Scripting.Dictionary")
    Set dicPosName = CreateObject("Scripting.Dictionary")
    
    Dim FSO As Object, TS As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim folderPath As String
    folderPath = ThisWorkbook.path & "\"
    
    Dim lineText As String
    Dim arr As Variant

    ' --- 1. 会社.ini（4カラム構成）の読み込み ---
    If FSO.FileExists(folderPath & "会社.ini") Then
        Set TS = FSO.OpenTextFile(folderPath & "会社.ini", 1)
        Do Until TS.AtEndOfStream
            lineText = TS.ReadLine
            If Trim(lineText) <> "" Then
                arr = Split(lineText, ",")
                If UBound(arr) >= 3 Then
                    ' ★復活：会社辞書 (会社コードを2桁にして会社名の前にくっつける「からくり」)
                    Dim kKey As Long: kKey = CLng(Trim(arr(0)))
                    If Not dicKaisyaName.Exists(kKey) Then
                        ' ここで「01.フジパングループ本社」のように合成します
                        dicKaisyaName.Add kKey, Format(kKey, "00") & "." & Trim(arr(1))
                    End If
                    
                    ' 工場辞書 (会社コード_工場コード -> 工場名)
                    Dim kojoKey As String: kojoKey = CLng(Trim(arr(0))) & "_" & CLng(Trim(arr(2)))
                    If Not dicKojoName.Exists(kojoKey) Then
                        dicKojoName.Add kojoKey, Trim(arr(3))
                    End If
                End If
            End If
        Loop
        TS.Close
    End If

    ' --- 2. 役職.ini（ハイブリッド構成）の読み込み ---
    If FSO.FileExists(folderPath & "役職.ini") Then
        Set TS = FSO.OpenTextFile(folderPath & "役職.ini", 1)
        Do Until TS.AtEndOfStream
            lineText = TS.ReadLine
            If Trim(lineText) <> "" Then
                arr = Split(lineText, ",")
                
                If UBound(arr) = 1 Then
                    ' 2カラム（職位, 役職名）
                    Dim pKey1 As String: pKey1 = Trim(arr(0))
                    If Not dicPosName.Exists(pKey1) Then
                        dicPosName.Add pKey1, Trim(arr(1))
                    End If
                ElseIf UBound(arr) = 2 Then
                    ' 3カラム（職位, 等級, 役職名）
                    Dim pKey2 As String: pKey2 = Trim(arr(0)) & "_" & Trim(arr(1))
                    If Not dicPosName.Exists(pKey2) Then
                        dicPosName.Add pKey2, Trim(arr(2))
                    End If
                End If
            End If
        Loop
        TS.Close
    End If

    Set TS = Nothing
    Set FSO = Nothing
End Sub

Private Sub subLoadIniToDic(ByVal filePath As String, ByRef dic1 As Object, Optional ByRef dic2 As Object = Nothing, Optional ByVal addNo As Boolean = False)
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim TS As Object: Dim line As String: Dim arr As Variant
    
    If Not FSO.FileExists(filePath) Then Exit Sub
    Set TS = FSO.OpenTextFile(filePath, 1, False)
    
    Do Until TS.AtEndOfStream
        line = TS.ReadLine
        If Trim(line) <> "" Then
            arr = Split(line, ",")
            
' --- 会社.ini の読み込み部分のみ抜粋 ---
    If Not dic2 Is Nothing Then
        If UBound(arr) >= 2 Then
            If IsNumeric(Trim(arr(0))) And IsNumeric(Trim(arr(2))) Then
                Dim cCode As String: cCode = Trim(arr(0)) ' 会社CD
                Dim fCode As String: fCode = Trim(arr(2)) ' 工場CD
                
                ' 会社名の辞書：キーは会社CD
                dic1(CLng(cCode)) = IIf(addNo, Format(cCode, "00") & "." & Trim(arr(1)), Trim(arr(1)))
                
                ' ★工場名の辞書：キーを「会社CD_工場CD」にする（例："1_199"）
                ' 別の会社の同じ工場番号と混ざらなくなる
                If UBound(arr) >= 3 Then
                    dic2(cCode & "_" & fCode) = Trim(arr(3))
                End If
            End If
        End If            ' B. その他
            Else
                If UBound(arr) >= 1 Then
                    If IsNumeric(Trim(arr(0))) Then
                        Dim key As Long: key = CLng(Trim(arr(0)))
                        dic1(key) = IIf(addNo, Format(key, "00") & "." & Trim(arr(1)), Trim(arr(1)))
                    End If
                End If
            End If
        End If
    Loop
    TS.Close
End Sub

Public Sub subExtractDataToSheet()
    Dim cn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim strSQL As String
    Dim ST As Worksheet: Set ST = ThisWorkbook.Worksheets("元データ")

    ' 1. 前回のデータをクリア（見出しは残す）
    ST.UsedRange.Offset(1).ClearContents

    ' 2. DB接続
    Set cn = CreateObject("ADODB.Connection")
    On Error GoTo Err_Conn
    cn.Open P_ConnectString ' modPublicで定義された接続文字列を使用

    ' 3. SQL文の構築
    ' 退職日（RTTDAT）を基準に抽出（等級・役職マスタ（MPDP01）は参照しない／旧プログラムと同等）
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM LIBIMF.IRTP01 "
    strSQL = strSQL & " WHERE RTDLT <> 'X' "
    ' RTKJNO=0 のレコードは本来存在しない（正常なデータではない）ため除外。
    ' 前任者は「メモリが飛んだのに書かれたレコード」と表現していた異常値。
    strSQL = strSQL & " AND RTKJNO <> 0 "

    ' 期間指定がある場合のみWHERE句を追加
    If P_DateFrom <> "" And P_DateTo <> "" Then
        strSQL = strSQL & " AND RTTDAT BETWEEN " & P_DateFrom & " AND " & P_DateTo
    End If

    ' 並び替え（退職日の新しい順）
    strSQL = strSQL & " ORDER BY RTTDAT DESC "
    
    ' 4. データ抽出実行
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, cn

    ' 5. シートへ貼り付け
    If Not rs.EOF Then
        ST.Range("A2").CopyFromRecordset rs
        
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            ' 抽出したデータのフィールド名（列名）を1行目に順番に書き出す
            ST.Cells(1, col + 1).Value = rs.Fields(col).Name
        Next col
    
    Else
        MsgBox "指定された期間のデータは見つかりませんでした。", vbInformation
    End If

    ' 6. 後片付け
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing
    Exit Sub

Err_Conn:
    MsgBox "DB接続エラー: " & Err.Description, vbCritical
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then Set cn = Nothing
End Sub

'② 加工とラベル付与(データ整形)
Public Sub subAddCalculatedColumns()
    Dim ST As Worksheet: Set ST = ThisWorkbook.Worksheets("元データ")
    Dim SS As Worksheet
    Dim i As Long, j As Long, h As Integer, writeRow As Long

    ' 集計用辞書
    Dim dictEmp As Object: Set dictEmp = CreateObject("Scripting.Dictionary")

    ' 1. 集計用シートの準備
    On Error Resume Next
    Set SS = ThisWorkbook.Worksheets("集計用")
    On Error GoTo 0
    If SS Is Nothing Then
        Set SS = ThisWorkbook.Worksheets.Add(After:=ST)
        SS.Name = "集計用"
    End If
    SS.Cells.Clear

' ★修正：会社だけでなく、役職辞書が空の時も確実に読み込む
    If dicKaisyaName Is Nothing Then Call subGetIniFileAll
    If dicKaisyaName.Count = 0 Then Call subGetIniFileAll
    If dicPosName Is Nothing Then Call subGetIniFileAll
    If dicPosName.Count = 0 Then Call subGetIniFileAll
    
    ' 2. ヘッダー作成
    Dim headers As Variant
    headers = Array("性別名", "勤続年数", "勤続区分", "会社名", "工場名", "役職名", "所属名", "退職年", "退職月", "区分", "集計理由項目", "社員番号", "元行", "人数フラグ")
    For h = 0 To UBound(headers)
        SS.Cells(1, h + 1).Value = headers(h)
    Next h

    ' 3. 列位置の特定
    Dim cSex As Long, cJoin As Long, cRetire As Long
    Dim cKai As Long, cKojo As Long, cPos As Long, cDept As Long, cEmp As Long
    Dim cDlt As Long
    On Error Resume Next
    cSex = ST.Rows(1).Find(What:="RTSEXC", LookAt:=xlWhole).Column
    cJoin = ST.Rows(1).Find(What:="RTNDAT", LookAt:=xlWhole).Column
    cRetire = ST.Rows(1).Find(What:="RTTDAT", LookAt:=xlWhole).Column
    cKai = ST.Rows(1).Find(What:="RTKAIC", LookAt:=xlWhole).Column
    cKojo = ST.Rows(1).Find(What:="RTKJNO", LookAt:=xlWhole).Column
    cPos = ST.Rows(1).Find(What:="RTSKUC", LookAt:=xlWhole).Column
    cDept = ST.Rows(1).Find(What:="RTSZBM", LookAt:=xlWhole).Column
    cEmp = ST.Rows(1).Find(What:="RTSYNO", LookAt:=xlWhole).Column
    cDlt = ST.Rows(1).Find(What:="RTDLT", LookAt:=xlWhole).Column
    On Error GoTo 0

    ' 4. 理由項目の定義
    Dim fieldsK As Variant: fieldsK = Array("RTKNIN", "RTKGYN", "RTKGYR", "RTKZAN", "RTKKYU", "RTKYAS", "RTKYAN", "RTKKIN", "RTKSYO", "RTKHYO", "RTKROU", "RTKKIT", "RTKKAI", "RTKKEK", "RTKSYU", "RTKIKU", "RTKKAG", "RTKKEN", "RTKHAI", "RTKRYU", "RTKSON")
    Dim namesK As Variant: namesK = Array("01.人間関係の悩み", "02.業務内容にやりがいを感じない", "03.業務量の多さ", "04.残業量の多さ", "05.給与", "06.年間休日が少ない", "07.休日が取れない", "08.勤務地", "09.昇格に不満", "10.評価に不満", "11.環境", "12.勤務時間が合わない", "13.会社の将来性に不安", "14.結婚", "15.出産", "16.育児", "17.介護", "18.健康面に不安がある", "19.配偶者転勤", "20.キャリアアップ", "21.その他")
    ' RTNROU（環境）は元データに存在しないため除外
    Dim fieldsN As Variant: fieldsN = Array("RTNNIN", "RTNGYN", "RTNSEI", "RTNGYR", "RTNZAN", "RTNKYU", "RTNYAS", "RTNKIN", "RTNJUU", "RTNSYO", "RTNHYO", "RTNKAI", "RTNSON")
    Dim namesN As Variant: namesN = Array("01.人間関係", "02.業務内容", "03.やりがい・成長", "04.業務量", "05.残業量", "06.給与", "07.年間休日の日数や希望", "08.勤務地", "09.柔軟な働き方", "10.昇格", "11.評価", "12.会社の将来性", "13.その他")

    ' 5. メインループ
    Dim lastRow As Long
    lastRow = ST.UsedRange.Rows(ST.UsedRange.Rows.Count).Row
    writeRow = 2
    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ' 完全な空行はスキップ
        If WorksheetFunction.CountA(ST.Rows(i)) = 0 Then GoTo NextPerson

        ' 削除フラグのチェック
        Dim valDlt As String: valDlt = ""
        If cDlt > 0 Then valDlt = Trim(CStr(ST.Cells(i, cDlt).Value))
        If valDlt <> "" Then GoTo NextPerson

        Dim valEmpRaw As String: valEmpRaw = ""
        If cEmp > 0 Then valEmpRaw = Trim(CStr(ST.Cells(i, cEmp).Value))

        Dim valSex As String: valSex = ""
        If cSex > 0 Then valSex = Trim(CStr(ST.Cells(i, cSex).Value))

        ' ★【修正1】不明1を消滅させる：性別が1(男性), 2(女性)以外の不正データはスキップ
        If valSex <> "1" And valSex <> "2" Then GoTo NextPerson

        ' ★【修正2】女性3名を復活させる：社員番号ではなく「行番号(イベント)」をキーにする！
        Dim eventKey As String: eventKey = "ROW_" & i

        ' --- 共通項目の取得と計算 ---
        Dim outCommon As Variant: ReDim outCommon(1 To 9)
        Dim valJoin As String: If cJoin > 0 Then valJoin = Trim(CStr(ST.Cells(i, cJoin).Value))
        Dim valRetire As String: If cRetire > 0 Then valRetire = Trim(CStr(ST.Cells(i, cRetire).Value))
        Dim kaiCode As Long: If cKai > 0 Then kaiCode = Val(ST.Cells(i, cKai).Value)
        Dim kojoCode As Long: If cKojo > 0 Then kojoCode = Val(ST.Cells(i, cKojo).Value)
        Dim valDept As String: If cDept > 0 Then valDept = ST.Cells(i, cDept).Value
        
        ' 職位・等級から役職名を判定（役職.iniの範囲指定・完全一致に対応）
        Dim codeP As String: If cPos > 0 Then codeP = Trim(CStr(ST.Cells(i, cPos).Value))
        Dim grade As String: grade = "" ' 等級列名に合わせて取得
        Dim cGrade As Long: cGrade = ST.Rows(1).Find(What:="RTTKYU", LookAt:=xlWhole).Column
        If cGrade > 0 Then grade = Trim(CStr(ST.Cells(i, cGrade).Value))

        ' デバッグ用：職位・等級・辞書件数を出力
        Debug.Print "codeP=" & codeP & ", grade=" & grade & ", dicPosName.Count=" & dicPosName.Count
        If dicPosName Is Nothing Then
            Debug.Print "dicPosName is Nothing"
        End If

        outCommon(6) = GetYakushokuName(codeP, grade, dicPosName)

        outCommon(1) = IIf(valSex = "1", "1.男性", "2.女性")

        Dim dIn As Date, dOut As Date, isValidDate As Boolean: isValidDate = False
        On Error Resume Next
        If IsNumeric(valJoin) And Len(valJoin) = 8 Then
            dIn = DateSerial(Left(valJoin, 4), Mid(valJoin, 5, 2), Right(valJoin, 2))
        ElseIf IsDate(valJoin) Then
            dIn = CDate(valJoin)
        End If
        If IsNumeric(valRetire) And Len(valRetire) = 8 Then
            dOut = DateSerial(Left(valRetire, 4), Mid(valRetire, 5, 2), Right(valRetire, 2))
        ElseIf IsDate(valRetire) Then
            dOut = CDate(valRetire)
        End If
        If dIn > 0 And dOut > 0 Then isValidDate = True
        On Error GoTo 0
        
        If isValidDate Then
            Dim yrs As Double: yrs = Round(DateDiff("d", dIn, dOut) / 365.25, 1)
            outCommon(2) = yrs
            outCommon(3) = Switch(yrs < 1, "01.1年未満", yrs < 3, "02.1-3年", yrs < 5, "03.3-5年", yrs < 10, "04.5-10年", True, "05.10年以上")
            outCommon(8) = Year(dOut) & "年"
            outCommon(9) = Month(dOut) & "月"
        End If

        ' 会社名・工場名のセット（辞書にない場合は空文字にする）
        If dicKaisyaName.Exists(kaiCode) Then
            outCommon(4) = dicKaisyaName(kaiCode)
        Else
            outCommon(4) = ""
        End If
        If dicKojoName.Exists(kaiCode & "_" & kojoCode) Then
            outCommon(5) = dicKojoName(kaiCode & "_" & kojoCode)
        Else
            outCommon(5) = ""
        End If        

        
        outCommon(7) = valDept
        

        ' --- (A) きっかけの展開 ---
        Dim hasKikkake As Boolean: hasKikkake = False
        For j = 0 To UBound(fieldsK)
            Dim colK As Long: colK = 0
            On Error Resume Next: colK = ST.Rows(1).Find(What:=fieldsK(j), LookAt:=xlWhole).Column: On Error GoTo 0
            If colK > 0 Then
                If Trim(CStr(ST.Cells(i, colK).Value)) = "1" Then
                    hasKikkake = True
                    Dim flagK As Integer: flagK = 0
                    If Not dictEmp.Exists(eventKey & "_きっかけ") Then
                        flagK = 1
                        dictEmp.Add eventKey & "_きっかけ", True
                    End If
                    Call subWriteRow(SS, writeRow, outCommon, "きっかけ", CStr(namesK(j)), valEmpRaw, i)
                    SS.Cells(writeRow, 14).Value = flagK
                    writeRow = writeRow + 1
                End If
            End If
        Next j
        
        If Not hasKikkake Then
            Dim flagK2 As Integer: flagK2 = 0
            If Not dictEmp.Exists(eventKey & "_きっかけ") Then
                flagK2 = 1
                dictEmp.Add eventKey & "_きっかけ", True
            End If
            Call subWriteRow(SS, writeRow, outCommon, "きっかけ", "未回答", valEmpRaw, i)
            SS.Cells(writeRow, 14).Value = flagK2
            writeRow = writeRow + 1
        End If

' --- (B) 期待の展開 ---
        Dim hasKitai As Boolean: hasKitai = False
        For j = 0 To UBound(fieldsN)
            Dim colN As Long: colN = 0
            On Error Resume Next: colN = ST.Rows(1).Find(What:=fieldsN(j), LookAt:=xlWhole).Column: On Error GoTo 0
            If colN > 0 Then
                If Trim(CStr(ST.Cells(i, colN).Value)) = "1" Then
                    hasKitai = True
                    Dim flagN As Integer: flagN = 0
                    ' 1人で複数の期待理由を出力できるようにキーを分ける
                    If Not dictEmp.Exists(eventKey & "_期待_" & j) Then
                        flagN = 1
                        dictEmp.Add eventKey & "_期待_" & j, True
                    End If
                    Call subWriteRow(SS, writeRow, outCommon, "期待", CStr(namesN(j)), valEmpRaw, i)
                    SS.Cells(writeRow, 14).Value = flagN
                    writeRow = writeRow + 1
                End If
            End If
        Next j
        
        If Not hasKitai Then
            Dim flagN2 As Integer: flagN2 = 0
            If Not dictEmp.Exists(eventKey & "_期待_未回答") Then
                flagN2 = 1
                dictEmp.Add eventKey & "_期待_未回答", True
            End If
            Call subWriteRow(SS, writeRow, outCommon, "期待", "未回答", valEmpRaw, i)
            SS.Cells(writeRow, 14).Value = flagN2
            writeRow = writeRow + 1
        End If
NextPerson:
    Next i

    SS.Columns.AutoFit
    Application.ScreenUpdating = True
End Sub

' --- 書き出し用サブ（引数の型とByRef/ByValを最適化） ---
Private Sub subWriteRow(ByVal ws As Worksheet, ByRef r As Long, ByRef common As Variant, ByVal kbn As String, ByVal reason As String, ByVal emp As String, ByVal rowIdx As Long)
    Dim c As Integer
    For c = 1 To 9
        ws.Cells(r, c).Value = common(c)
    Next c
    ws.Cells(r, 10).Value = kbn
    ws.Cells(r, 11).Value = reason
    ws.Cells(r, 12).Value = emp
    ws.Cells(r, 13).Value = rowIdx
End Sub

' --- ダッシュボードを初期化 ---
Public Sub subResetDashboard()
    Dim ST As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    
    On Error Resume Next
    Set ST = ThisWorkbook.Worksheets("分析グラフ")
    On Error GoTo 0
    If ST Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' 1. グラフを削除
    On Error Resume Next
    ST.ChartObjects.Delete
    On Error GoTo 0

    ' 2. 2行目の条件をクリア
    ST.Range("B2").Value = ""
    ST.Range("F2").Value = ""
    ST.Range("J2").Value = ""

    ' 3. ピボットテーブルの中身だけを空にする
    On Error Resume Next
    For Each pt In ST.PivotTables
        pt.ManualUpdate = True
        
        ' 表示されているすべてのフィールドを非表示
        For Each pf In pt.DataFields: pf.Orientation = xlHidden: Next pf
        For Each pf In pt.RowFields: pf.Orientation = xlHidden: Next pf
        For Each pf In pt.ColumnFields: pf.Orientation = xlHidden: Next pf
        For Each pf In pt.PageFields: pf.Orientation = xlHidden: Next pf
        
        pt.ManualUpdate = False
    Next pt
    On Error GoTo 0

    ' 完了後のカーソル位置
    ST.Activate
    ST.Range("A1").Select

    Application.ScreenUpdating = True
End Sub

' --- ダッシュボード終了ボタン（保存なし） ---
Public Sub subCloseSafely()
    Dim wb As Workbook
    Dim visibleCount As Integer
    
    visibleCount = 0
    For Each wb In Workbooks
        If wb.Windows.Count > 0 Then
            If wb.Windows(1).Visible = True Then
                visibleCount = visibleCount + 1
            End If
        End If
    Next wb
    
    Application.DisplayAlerts = False
    
    If visibleCount <= 1 Then
        Application.Quit
        ThisWorkbook.Close SaveChanges:=False ' Quitの実行保証
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

' 職位・等級から役職名を判定（範囲指定・完全一致対応）
Function GetYakushokuName(ByVal codeP As String, ByVal grade As String, ByVal dicPosName As Object) As String
    Dim keyFull As String
    Dim k As Variant
    Dim iniArr As Variant
    Dim pos As String, cond As String, name As String
    Dim g As Long, gVal As Long

    ' 1. 完全一致（職位,等級）優先
    keyFull = codeP & "_" & grade
    If dicPosName.Exists(keyFull) Then
        GetYakushokuName = dicPosName(keyFull)
        Exit Function
    End If

    ' 2. 範囲指定（例: 40,<=21,07.一般）をiniから探す
    For Each k In dicPosName.Keys
        iniArr = Split(k, "_")
        If UBound(iniArr) = 1 Then
            pos = iniArr(0)
            cond = iniArr(1)
            If pos = codeP Then
                ' 範囲指定パターン
                If Left(cond, 2) = "<=" Then
                    gVal = Val(Mid(cond, 3))
                    If IsNumeric(grade) And Val(grade) <= gVal Then
                        GetYakushokuName = dicPosName(k)
                        Exit Function
                    End If
                ElseIf Left(cond, 2) = ">=" Then
                    gVal = Val(Mid(cond, 3))
                    If IsNumeric(grade) And Val(grade) >= gVal Then
                        GetYakushokuName = dicPosName(k)
                        Exit Function
                    End If
                ElseIf cond = "*" Then
                    GetYakushokuName = dicPosName(k)
                    Exit Function
                End If
            End If
        End If
    Next k

    ' 3. 職位のみ一致
    If dicPosName.Exists(codeP) Then
        GetYakushokuName = dicPosName(codeP)
        Exit Function
    End If

    ' 4. どれにも該当しない場合はコードを返す
    GetYakushokuName = codeP
End Function