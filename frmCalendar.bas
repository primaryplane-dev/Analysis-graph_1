Option Explicit

Private Sub UserForm_Initialize()
    Dim dateWK      As Date
    
    P_Calendar_FLG = False
    dateWK = Now
    If P_DATEC > 0 Then dateWK = P_DATEC
    Me.txtDate = Format(dateWK, "yyyy/mm/dd")
    
    'カレンダを描画する
    Call SetCalender(dateWK)
End Sub

'決定ボタン押下時
Private Sub cmdOK_Click()
    Call subOK
End Sub

Private Sub subOK()
    P_DATEC = CDate(Me.txtDate)
    P_Calendar_FLG = True
    Unload Me
End Sub

'キャンセルボタン押下時
Private Sub cmdCancel_Click()
    Unload Me
End Sub


'--- ↓日付描画関連↓ ---
'カレンダ描画
Private Sub SetCalender(ByVal i_Date As Date)
    Dim i           As Integer
    Dim dateWK      As Date
    Dim dateFR      As Date
    Dim dateTO      As Date
    Dim strName     As String

    Me.txtYear = Format(i_Date, "yyyy")
    Me.txtMonth = Format(i_Date, "m")
    
    'カレンダをクリアする
    For i = 1 To 42
        strName = "cln" & Format(i, "00")
        Me(strName).Caption = ""
        Me(strName).Tag = 0
        Me(strName).Enabled = False
        Me(strName).ForeColor = vbBlack
    Next
    '月初～月末の日付を入れる
    dateFR = DateSerial(Year(i_Date), Month(i_Date), 1)             '月初
    dateTO = DateSerial(Year(i_Date), Month(i_Date) + 1, 1) - 1     '月末
    i = Format(dateFR, "w") - 1                                     'w:曜日の番号(1～7)
    For dateWK = dateFR To dateTO
        i = i + 1
        strName = "cln" & Format(i, "00")
        Me(strName).Caption = Format(dateWK, "d")
        Me(strName).Tag = Format(dateWK, "yyyy/mm/dd")
        Me(strName).Enabled = True
        If Weekday(Me(strName).Tag) = vbSunday Then Me(strName).ForeColor = vbRed
        If Me(strName).Tag = Me.txtDate Then Me(strName).ForeColor = vbBlue
    Next
End Sub

' < ボタン押下時
Private Sub cmdMonthDown_Click()
    Dim dateWK      As Date
    
    'カレンダを描画する(前月)
    cmdMonthDown.Enabled = False
    dateWK = DateSerial(Me.txtYear, Me.txtMonth, 1)
    dateWK = DateAdd("m", -1, dateWK)
    Call SetCalender(dateWK)
    cmdMonthDown.Enabled = True
End Sub

' > ボタン押下時
Private Sub cmdMonthUp_Click()
    Dim dateWK      As Date
    
    'カレンダを描画する(翌月)
    cmdMonthUp.Enabled = False
    dateWK = DateSerial(Me.txtYear, Me.txtMonth, 1)
    dateWK = DateAdd("m", 1, dateWK)
    Call SetCalender(dateWK)
    cmdMonthUp.Enabled = True
End Sub


'--- ↓日付選択関連↓ ---
'今日ボタン押下時
Private Sub cmdToday_Click()
    Me.txtDate = Format(Now, "yyyy/mm/dd")
    Call SetCalender(Now)
End Sub

Private Sub SetDate(ByVal i_ctlIdx As Integer)
    Dim i           As Integer
    Dim strName     As String
    
    'カレンダ部分の色を初期表示にもどす
    For i = 1 To 42
        strName = "cln" & Format(i, "00")
        Me(strName).ForeColor = vbBlack
        If Weekday(Me(strName).Tag) = vbSunday Then Me(strName).ForeColor = vbRed
    Next
    
    '選択された日を青文字表示
    strName = "cln" & Format(i_ctlIdx, "00")
    Me(strName).ForeColor = vbBlue
    Me.txtDate = Me(strName).Tag
End Sub

Private Sub cln01_Enter()
    Call SetDate(1)
End Sub

Private Sub cln02_Enter()
    Call SetDate(2)
End Sub

Private Sub cln03_Enter()
    Call SetDate(3)
End Sub

Private Sub cln04_Enter()
    Call SetDate(4)
End Sub

Private Sub cln05_Enter()
    Call SetDate(5)
End Sub

Private Sub cln06_Enter()
    Call SetDate(6)
End Sub

Private Sub cln07_Enter()
    Call SetDate(7)
End Sub

Private Sub cln08_Enter()
    Call SetDate(8)
End Sub

Private Sub cln09_Enter()
    Call SetDate(9)
End Sub

Private Sub cln10_Enter()
    Call SetDate(10)
End Sub

Private Sub cln11_Enter()
    Call SetDate(11)
End Sub

Private Sub cln12_Enter()
    Call SetDate(12)
End Sub

Private Sub cln13_Enter()
    Call SetDate(13)
End Sub

Private Sub cln14_Enter()
    Call SetDate(14)
End Sub

Private Sub cln15_Enter()
    Call SetDate(15)
End Sub

Private Sub cln16_Enter()
    Call SetDate(16)
End Sub

Private Sub cln17_Enter()
    Call SetDate(17)
End Sub

Private Sub cln18_Enter()
    Call SetDate(18)
End Sub

Private Sub cln19_Enter()
    Call SetDate(19)
End Sub

Private Sub cln20_Enter()
    Call SetDate(20)
End Sub

Private Sub cln21_Enter()
    Call SetDate(21)
End Sub

Private Sub cln22_Enter()
    Call SetDate(22)
End Sub

Private Sub cln23_Enter()
    Call SetDate(23)
End Sub

Private Sub cln24_Enter()
    Call SetDate(24)
End Sub

Private Sub cln25_Enter()
    Call SetDate(25)
End Sub

Private Sub cln26_Enter()
    Call SetDate(26)
End Sub

Private Sub cln27_Enter()
    Call SetDate(27)
End Sub

Private Sub cln28_Enter()
    Call SetDate(28)
End Sub

Private Sub cln29_Enter()
    Call SetDate(29)
End Sub

Private Sub cln30_Enter()
    Call SetDate(30)
End Sub

Private Sub cln31_Enter()
    Call SetDate(31)
End Sub

Private Sub cln32_Enter()
    Call SetDate(32)
End Sub

Private Sub cln33_Enter()
    Call SetDate(33)
End Sub

Private Sub cln34_Enter()
    Call SetDate(34)
End Sub

Private Sub cln35_Enter()
    Call SetDate(35)
End Sub

Private Sub cln36_Enter()
    Call SetDate(36)
End Sub

Private Sub cln37_Enter()
    Call SetDate(37)
End Sub

Private Sub cln38_Enter()
    Call SetDate(38)
End Sub

Private Sub cln39_Enter()
    Call SetDate(39)
End Sub

Private Sub cln40_Enter()
    Call SetDate(40)
End Sub

Private Sub cln41_Enter()
    Call SetDate(41)
End Sub

Private Sub cln42_Enter()
    Call SetDate(42)
End Sub
