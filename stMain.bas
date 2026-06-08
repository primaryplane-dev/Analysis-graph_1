Option Explicit

Sub subOpenSelect()
    Dim obj As New frmSelect
    obj.Show

    If P_Regist = True Then
        Call subMain
    End If

    Set obj = Nothing
End Sub
'
Public Sub HandlePivotTableUpdate(ByVal ws As Worksheet, ByVal Target As PivotTable)
    Dim pc As Chart
    Dim targetName As String

    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    On Error Resume Next
    targetName = Target.Name
    On Error GoTo 0

    Application.ScreenUpdating = False
    WritePivotEventLog ws, "[Event] PivotTableUpdate target=" & targetName

    ' 1. リフレッシュ処理
    On Error Resume Next
'    Call subRefreshGraphCleanly
    Call subTryApplyRoleReasonDetailFromSelection
    On Error GoTo 0

    ' 2. グラフタイトル更新（今の分析名を表示）
    On Error Resume Next
    Set pc = ws.ChartObjects("退職分析グラフ").Chart
    If Not pc Is Nothing Then
        pc.ChartTitle.Text = fnTrimCode(P_Pattern)
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub

Public Sub WritePivotEventLog(ByVal ws As Worksheet, ByVal msg As String)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Range("AA2").Value = "EventDebug"
    ws.Range("AA3").Value = msg
    On Error GoTo 0
End Sub
