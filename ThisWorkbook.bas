Option Explicit

Private Sub Workbook_Open()
    Application.EnableEvents = True

    ' ダッシュボード初期化（分析グラフ）
    Call subResetDashboard
End Sub

Private Sub Workbook_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
    Dim ws As Worksheet
    Dim pc As Chart

    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If CStr(Sh.Name) <> "分析グラフ" Then Exit Sub
    Set ws = Sh

    If P_IsPivotBuilding Then
        WritePivotEventLog ws, "[Event] skip during build"
        Exit Sub
    End If

    WritePivotEventLog ws, "[Event] Workbook_SheetPivotTableUpdate fired target=" & Target.Name

    Application.ScreenUpdating = False

    Call subTryApplyRoleReasonDetailFromSelection

    Set pc = ws.ChartObjects("退職分析グラフ").Chart
    If Not pc Is Nothing Then
        pc.ChartTitle.Text = fnTrimCode(P_Pattern)
    End If

    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub WritePivotEventLog(ByVal ws As Worksheet, ByVal msg As String)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Range("AA2").Value = "EventDebug"
    ws.Range("AA3").Value = msg
    On Error GoTo 0
End Sub

