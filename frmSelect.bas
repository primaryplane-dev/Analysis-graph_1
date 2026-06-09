Option Explicit

' --- フォーム初期化 ---
Private Sub UserForm_Initialize()
    Dim lastPatternIndex As Long

    ' 1. 期間指定の初期設定（チェックなし・入力不可）
    Me.chkDate.Value = False
    Call subControlDateInput(False)
    
    ' 2. コンボボックスに13パターンを登録
    With Me.cmbPattern
        .AddItem "01.男女比 (全体)"
        .AddItem "02.勤続年数比 (全体)"
        .AddItem "03.男女別 × 理由"
        .AddItem "04.勤続年数別 × 理由"
        .AddItem "05.役職別 × 理由"
        .AddItem "06.会社別 × 理由"
        .AddItem "07.工場別 × 理由"
        .AddItem "08.所属別 × 理由"
        .AddItem "09.月別 × 退職者推移"
        .AddItem "10.年別 × 退職者推移"
        .AddItem "11.1年間の月別推移 (理由)"
        .AddItem "12.3年間の年別推移 (理由)"
'        .AddItem "13.回答内容をエクセル出力"

        lastPatternIndex = fnFindPatternIndex(P_Pattern)
        If lastPatternIndex >= 0 Then
            .ListIndex = lastPatternIndex
        Else
            .ListIndex = 0 ' デフォルトで一番上を選択
        End If
    End With
    
    ' 3. 集計対象の初期選択
    If P_Kubun2 = 2 Then
        Me.optNext.Value = True
    Else
        Me.optKikkake.Value = True
    End If

    ' 4. パターンに応じて条件区分の活性を切り替え
    Call subControlKubunByPattern
End Sub

Private Function fnFindPatternIndex(ByVal patternText As String) As Long
    Dim i As Long

    fnFindPatternIndex = -1
    If Trim$(patternText) = "" Then Exit Function

    For i = 0 To Me.cmbPattern.ListCount - 1
        If StrComp(CStr(Me.cmbPattern.List(i)), patternText, vbTextCompare) = 0 Then
            fnFindPatternIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub cmbPattern_Change()
    Call subControlKubunByPattern
End Sub

' --- 期間指定チェックボックスの制御 ---
Private Sub chkDate_Change()
    If Me.chkDate.Value = True Then
        ' チェックされたらデフォルト期間（本日～1年前）をセット
        Me.txtDateT.Value = Format(Date, "yyyy/mm/dd")
        Me.txtDateF.Value = Format(DateAdd("yyyy", -1, Date), "yyyy/mm/dd")
        Call subControlDateInput(True)
    Else
        ' チェックが外れたら日付をクリア
        Me.txtDateF.Value = ""
        Me.txtDateT.Value = ""
        Call subControlDateInput(False)
    End If
End Sub

' 日付入力関連の活性・非活性を切り替えるサブ処理
Private Sub subControlDateInput(ByVal IsEnabled As Boolean)
    Me.txtDateF.Enabled = IsEnabled
    Me.txtDateT.Enabled = IsEnabled
    Me.btnCalF.Enabled = IsEnabled ' カレンダー呼び出しボタン(開始)
    Me.btnCalT.Enabled = IsEnabled ' カレンダー呼び出しボタン(終了)
End Sub

' パターン1・2は条件区分を使用しないため非活性化
Private Sub subControlKubunByPattern()
    Dim pNum As Integer
    Dim disableKubun As Boolean

    pNum = Val(Left$(Me.cmbPattern.Value, 2))
    disableKubun = (pNum = 1 Or pNum = 2 Or (pNum >= 9 And pNum <= 12))

    Me.optKikkake.Enabled = Not disableKubun
    Me.optNext.Enabled = Not disableKubun

    If disableKubun Then
        Me.optKikkake.Value = True
    End If
End Sub

' --- カレンダーボタン押下（開始日） ---
Private Sub btnCalF_Click()
    Dim d As Date
    If IsDate(Me.txtDateF.Value) Then
        d = CDate(Me.txtDateF.Value)
    Else
        d = Date
    End If
    Me.txtDateF.Value = fncGetCalendarDate(d)
End Sub

' --- カレンダーボタン押下（終了日） ---
Private Sub btnCalT_Click()
    Dim d As Date
    
    ' IIfは使わず、If文で安全に日付を取得する
    If IsDate(Me.txtDateT.Value) Then
        d = CDate(Me.txtDateT.Value)
    Else
        d = Date ' 空欄なら今日の日付をデフォルトにする
    End If
    
    ' カレンダーを呼び出す
    Me.txtDateT.Value = fncGetCalendarDate(d)
End Sub

' カレンダーフォームを呼び出して結果を返す共通関数
Private Function fncGetCalendarDate(ByVal initDate As Date) As String
    Dim f As New frmCalendar
    P_DATEC = initDate ' 共有変数へ初期値を渡す
    P_Calendar_FLG = False
    
    f.Show ' カレンダーをモーダル表示
    
    If P_Calendar_FLG Then
        fncGetCalendarDate = Format(P_DATEC, "yyyy/mm/dd")
    Else
        fncGetCalendarDate = "" ' キャンセル時は空を返す
    End If
    Set f = Nothing
End Function

' --- 実行ボタン ---
Private Sub cmdExecute_Click()
    ' 期間指定ありの場合の入力チェック
    If Me.chkDate.Value = True Then
        If Not IsDate(Me.txtDateF.Value) Or Not IsDate(Me.txtDateT.Value) Then
            MsgBox "期間が正しく入力されていません。", vbExclamation
            Exit Sub
        End If
    End If

    ' 公用変数に値をセットして処理へ渡す
    P_DateFrom = IIf(Me.chkDate.Value, Format(Me.txtDateF.Value, "yyyymmdd"), "")
    P_DateTo = IIf(Me.chkDate.Value, Format(Me.txtDateT.Value, "yyyymmdd"), "")
    If Val(Left$(Me.cmbPattern.Value, 2)) <= 2 Or (Val(Left$(Me.cmbPattern.Value, 2)) >= 9 And Val(Left$(Me.cmbPattern.Value, 2)) <= 12) Then
        P_Kubun2 = 1
    Else
        P_Kubun2 = IIf(Me.optKikkake.Value, 1, 2)
    End If
    P_Pattern = Me.cmbPattern.Value ' 選択された13パターンの名前を保存
    
    P_Regist = True
    Me.Hide
End Sub

' --- キャンセルボタン ---
Private Sub cmdCancel_Click()
    P_Regist = False
    Unload Me
End Sub

Private Sub subOpenCalendar()
    Dim obj As New frmCalendar
    obj.Show
    Set obj = Nothing
End Sub


