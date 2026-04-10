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
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim ws As Worksheet: Set ws = Me
    Dim pc As Chart

    Application.ScreenUpdating = False

    ' 1. リフレッシュ処理
    On Error Resume Next
'    Call subRefreshGraphCleanly
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
