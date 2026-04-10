Option Explicit

Public Sub subUpdatePivotByMenu(ByVal selectedPattern As String)
    Dim ST As Worksheet, SS As Worksheet
    Dim pt As PivotTable, pc As Chart, pcCache As PivotCache
    Dim lastRowSS As Long, pNum As Integer
    
    Application.ScreenUpdating = False
    
    Set ST = ThisWorkbook.Worksheets("分析グラフ")
    Set SS = ThisWorkbook.Worksheets("集計用")
    
    ST.Activate
    ActiveWindow.FreezePanes = False
    
    lastRowSS = SS.Cells(SS.Rows.Count, 1).End(xlUp).Row
    If lastRowSS < 2 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- 列幅固定 ---
    ST.Columns("B").ColumnWidth = 31
    ST.Columns("C:AZ").ColumnWidth = 8.5
    ST.Columns("A").ColumnWidth = 2

    ' --- 画面クリア ---
    On Error Resume Next
    ST.ChartObjects.Delete
    For Each pt In ST.PivotTables: pt.TableRange2.Clear: Next pt
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches: sc.Delete: Next sc
    ST.Range("B3:X1000").Clear
    ST.Rows("1:2").ClearContents
    ST.Rows("1:2").ClearFormats
    On Error GoTo 0

    ' --- タイトル＆条件設定 ---
    With ST
        With .Range("A1:Z2")
            .Interior.Color = RGB(141, 180, 226)
            .Font.Color = RGB(0, 32, 96)
        End With
        
        .Range("B1").Value = "ジョブリターン分析"
        .Range("B1").Font.Size = 18
        .Range("B1").Font.Bold = True
        
        .Range("C1").Value = "現在表示中"
        .Range("C1").Font.Size = 12
        .Range("C1").Font.Size = 12
        .Range("C2").Value = fnTrimCode(selectedPattern)
        
        .Range("G1").Value = "期間"
        .Range("G1").Font.Size = 12
        If P_DateFrom <> "" And P_DateTo <> "" Then
            .Range("G2").Value = Format(P_DateFrom, "@@@@/@@/@@") & " ～ " & Format(P_DateTo, "@@@@/@@/@@")
        Else
            .Range("G2").Value = "全期間"
        End If
        
        .Range("K1").Value = "区分"
        .Range("K1").Font.Size = 12
        .Range("K2").Value = IIf(P_Kubun2 = 2, "次のステップへの期待", "退職を考えたきっかけ")
    End With

    ' --- ピボットテーブル作成（100行目から） ---
    Set pcCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SS.Range("A1").CurrentRegion)
    Set pt = pcCache.CreatePivotTable(TableDestination:=ST.Range("B100"), TableName:="退職分析ピボット")
    pt.SaveData = True
    pt.ManualUpdate = True ' ★計算ストップ
    pt.HasAutoFormat = False
    pt.PreserveFormatting = True

    ' --- パターン別 値エリアの切り替え ---
    pNum = Val(Left(selectedPattern, 2))
    
    Select Case pNum
        Case 1, 2, 5, 6, 7, 8, 9, 10
            With pt.PivotFields("人数フラグ")
                .Orientation = xlDataField
                .Function = xlSum
                .Name = "退職者数 "
                .NumberFormat = "#,##0" ' ★「人」を削除
            End With
        Case 3, 4, 11, 12
            With pt.PivotFields("社員番号")
                .Orientation = xlDataField
                .Function = xlCount
                .Name = "理由延べ数 "
                .NumberFormat = "#,##0" ' ★「件」を削除
            End With
    End Select

    With pt.PivotFields("区分")
        .Orientation = xlPageField
        .Position = 1
        On Error Resume Next
        .CurrentPage = IIf(P_Kubun2 = 2, "期待", "きっかけ")
        On Error GoTo 0
    End With

    Dim axisField As String: axisField = ""
    Select Case pNum
        Case 1, 3: axisField = "性別名"
        Case 2, 4: axisField = "勤続区分"
        Case 5: axisField = "役職名"
        Case 6: axisField = "会社名"
        Case 7: axisField = "工場名"
        Case 8: axisField = "所属名"
        Case 9, 11: axisField = "退職月"
        Case 10, 12: axisField = "退職年"
    End Select
    If axisField <> "" Then pt.PivotFields(axisField).Orientation = xlRowField
    
    If pNum >= 3 And pNum <= 8 Then
        pt.PivotFields("集計理由項目").Orientation = xlRowField
    ElseIf pNum = 11 Or pNum = 12 Then
        pt.PivotFields("集計理由項目").Orientation = xlColumnField
    End If

    ' --- グラフ作成 ---
    Dim cho As ChartObject
    Set cho = ST.ChartObjects.Add(Left:=ST.Range("B4").Left, Top:=ST.Range("B4").Top, Width:=730, Height:=600)
    cho.Name = "退職分析グラフ"
    Set pc = cho.Chart
    pc.SetSourceData Source:=pt.TableRange1
    
    Select Case pNum
        Case 1, 2, 3, 4, 5, 6, 7, 8: pc.ChartType = xlPie
        Case 11, 12: pc.ChartType = xlColumnStacked
        Case Else: pc.ChartType = xlColumnClustered
    End Select
    
    ' --- スライサー作成 ---
    Dim slNames As Variant, slItem As Variant
    Dim i As Integer: i = 0
    slNames = Array("会社名", "工場名", "勤続区分", "性別名", "役職名")
    
    For Each slItem In slNames
        Set sc = Nothing
        On Error Resume Next
        Set sc = ThisWorkbook.SlicerCaches.Add(pt, CStr(slItem))
        On Error GoTo 0
        
        If Not sc Is Nothing Then
            sc.Slicers.Add ST, Name:="Sl_" & slItem, Caption:=CStr(slItem), _
                           Top:=ST.Range("M4").Top, _
                           Left:=ST.Range("M4").Left + (i * 150), _
                           Width:=140, Height:=380
            i = i + 1
        End If
    Next slItem
    
    ' --- 完了処理 ---
    ' 計算再開（グラフ作られる）
    pt.ManualUpdate = False
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    ' 図形が作られたあとにタイトルと値をセット
    pc.HasTitle = True
    pc.ChartTitle.Text = fnTrimCode(selectedPattern)
    
    On Error Resume Next
    pc.ApplyDataLabels Type:=xlDataLabelsShowValue
    
    ' 凡例を下部に配置して折り返させる
    pc.HasLegend = True
    pc.Legend.Position = xlLegendPositionBottom
    
    
    Dim srs As Series
    For Each srs In pc.SeriesCollection
        srs.HasDataLabels = True
        srs.DataLabels.ShowValue = True
    Next srs
    On Error GoTo 0
    
    ' カーソルを左上に戻す
    ST.Activate
    ST.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    Application.ScreenUpdating = True
End Sub

