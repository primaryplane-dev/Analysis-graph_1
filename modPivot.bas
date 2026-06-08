Option Explicit

Private Const MAX_REASON_ITEMS As Long = 8
Private Const DETAIL_PATTERN_ROLE_REASON As Long = 5
Private Const MAIN_PIVOT_NAME As String = "退職分析ピボット"
Private Const MAIN_CHART_NAME As String = "退職分析グラフ"
Private Const DETAIL_PIVOT_NAME As String = "退職分析ピボット_詳細"
Private Const DETAIL_CHART_NAME As String = "退職分析グラフ_詳細"
Private m_IsApplyingRoleDetail As Boolean
Private m_DetailAxisField As String
Private m_DetailAxisValue As String
Private m_DebugRow As Long

Public Sub subUpdatePivotByMenu(ByVal selectedPattern As String)
    Dim ST As Worksheet, SS As Worksheet
    Dim pt As PivotTable, pc As Chart, pcCache As PivotCache
    Dim detailPt As PivotTable
    Dim lastRowSS As Long, pNum As Integer
    Dim axisField As String
    Dim isTwoStepReason As Boolean

    P_IsPivotBuilding = True
    Application.EnableEvents = True
    Application.ScreenUpdating = False

    Set ST = ThisWorkbook.Worksheets("分析グラフ")
    Set SS = ThisWorkbook.Worksheets("集計用")

    ST.Activate
    ActiveWindow.FreezePanes = False

    lastRowSS = SS.Cells(SS.Rows.Count, 1).End(xlUp).Row
    If lastRowSS < 2 Then
        Application.ScreenUpdating = True
        P_IsPivotBuilding = False
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
    Set pt = pcCache.CreatePivotTable(TableDestination:=ST.Range("B100"), TableName:=MAIN_PIVOT_NAME)
    pt.SaveData = True
    pt.ManualUpdate = True ' ★計算ストップ
    pt.HasAutoFormat = False
    pt.PreserveFormatting = True

    ' --- パターン別 値エリアの切り替え ---
    pNum = Val(Left(selectedPattern, 2))
    P_SelectedRoleDetail = ""
    m_DetailAxisField = ""
    m_DetailAxisValue = ""

    Select Case pNum
        Case 1, 2, 5, 6, 7, 8, 9, 10
            With pt.PivotFields("人数フラグ")
                .Orientation = xlDataField
                .Function = xlSum
                .name = "退職者数 "
                .NumberFormat = "#,##0" ' ★「人」を削除
            End With
        Case 3, 4, 11, 12
            With pt.PivotFields("社員番号")
                .Orientation = xlDataField
                .Function = xlCount
                .name = "理由延べ数 "
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

    axisField = GetAxisFieldByPattern(pNum)
    If axisField <> "" Then pt.PivotFields(axisField).Orientation = xlRowField

    isTwoStepReason = ShouldUseTwoStepReason(pt, pNum)

    If pNum >= 3 And pNum <= 8 And pNum <> DETAIL_PATTERN_ROLE_REASON And Not isTwoStepReason Then
        pt.PivotFields("集計理由項目").Orientation = xlRowField
        ApplyTopNToReason pt, MAX_REASON_ITEMS
    ElseIf pNum = 11 Or pNum = 12 Then
        pt.PivotFields("集計理由項目").Orientation = xlColumnField
    End If

    ' --- グラフ作成 ---
    Dim cho As ChartObject
    Set cho = ST.ChartObjects.Add(Left:=ST.Range("B4").Left, Top:=ST.Range("B4").Top, Width:=730, Height:=600)
    cho.name = MAIN_CHART_NAME
    Set pc = cho.Chart
    pc.SetSourceData Source:=pt.TableRange1

    Select Case pNum
        Case 1, 2, DETAIL_PATTERN_ROLE_REASON: pc.ChartType = xlPie
        Case 3, 4, 5, 6, 7, 8: pc.ChartType = xlBarClustered
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
            sc.Slicers.Add ST, name:="Sl_" & slItem, Caption:=CStr(slItem), _
                           Top:=ST.Range("M4").Top, _
                           Left:=ST.Range("M4").Left + (i * 150), _
                           Width:=140, Height:=380
            i = i + 1
        End If
    Next slItem

    ' スライサーは2段階目用ピボットに接続し、1段階目を非連動にする
    Set detailPt = GetOrCreateDetailPivot(ST, pcCache)
    If Not detailPt Is Nothing Then
        RebindSlicersToDetailPivot pt, detailPt
    End If

    ' --- 完了処理 ---
    ' 計算再開（グラフ作られる）
    pt.ManualUpdate = False

    ActiveWorkbook.ShowPivotTableFieldList = False

    ' 図形が作られたあとにタイトルと値をセット
    pc.HasTitle = True
    pc.ChartTitle.Text = fnBuildChartTitle()

    On Error Resume Next
    pc.ApplyDataLabels Type:=xlDataLabelsShowValue

    ' 凡例は常に表示
    pc.HasLegend = True
    pc.Legend.Position = xlLegendPositionRight


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

    ' パターン3～8は初期表示時にも2段階目を作成する
    If pNum >= 3 And pNum <= 8 Then
        subTryApplyRoleReasonDetailFromSelection
    End If

    P_IsPivotBuilding = False
End Sub

Public Sub subTryApplyRoleReasonDetailFromSelection()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pc As Chart
    Dim mainCho As ChartObject
    Dim selectedAxisItem As String
    Dim axisField As String
    Dim pNum As Integer
    Dim hasDetailChart As Boolean
    Dim sourcePt As PivotTable
    Dim pivotAxisItem As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("分析グラフ")
    If ws Is Nothing Then Exit Sub
    If ws.PivotTables.Count = 0 Then Exit Sub
    Set pt = ws.PivotTables(MAIN_PIVOT_NAME)
    Set mainCho = ws.ChartObjects(MAIN_CHART_NAME)
    Set pc = mainCho.Chart
    On Error GoTo 0

    hasDetailChart = DetailChartExists(ws)

    If pt Is Nothing Or pc Is Nothing Then Exit Sub
    If m_IsApplyingRoleDetail Then
        WriteDetailDebugLog "[Detail] skip reentry"
        Exit Sub
    End If

    Set sourcePt = GetOrCreateDetailPivot(ws, pt.PivotCache)
    If sourcePt Is Nothing Then Set sourcePt = pt

    pNum = Val(Left(P_Pattern, 2))
    axisField = GetAxisFieldByPattern(pNum)
    If axisField = "" Then
        axisField = m_DetailAxisField
        If axisField <> "" Then
            pNum = DETAIL_PATTERN_ROLE_REASON
        End If
    End If

    If axisField = "" Then
        WriteDetailDebugLog "[Detail] skip axisField unresolved pattern=" & P_Pattern
        RemoveDetailArtifacts ws
        Exit Sub
    End If

    ResetDetailDebugLog
    WriteDetailDebugLog "[Detail] pattern=" & P_Pattern & " pNum=" & CStr(pNum) & " axisField=" & axisField
    pivotAxisItem = GetSingleSelectedRole(sourcePt, axisField)
    WriteDetailDebugLog "[Detail] axisField=" & axisField & " pivotSelected=" & pivotAxisItem

    selectedAxisItem = pivotAxisItem
    If selectedAxisItem = "" Then
        selectedAxisItem = GetSingleSelectedAxisFromSlicer(axisField)
        WriteDetailDebugLog "[Detail] axisField=" & axisField & " slicerSelected=" & selectedAxisItem
    Else
        WriteDetailDebugLog "[Detail] axisField=" & axisField & " slicerSelected=(skip: pivot priority)"
    End If

    If selectedAxisItem = "" Then
        selectedAxisItem = GetSingleSelectedRole(sourcePt, axisField)
        WriteDetailDebugLog "[Detail] axis fallback(GetSingleSelectedRole)=" & selectedAxisItem
    End If

    m_IsApplyingRoleDetail = True
    On Error GoTo SafeExit
    ApplyRoleReasonDetail ws, pt, mainCho, axisField, selectedAxisItem, pNum

SafeExit:
    If Err.Number <> 0 Then
        WriteDetailDebugLog "[Detail] subTry error=" & CStr(Err.Number) & " " & Err.Description
        Err.Clear
    End If
    m_IsApplyingRoleDetail = False
End Sub

Private Function HasAxisSlicerForField(ByVal fieldName As String) As Boolean
    Dim sc As SlicerCache

    HasAxisSlicerForField = False
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        If IsAxisSlicerCache(sc, fieldName) Then
            HasAxisSlicerForField = True
            Exit For
        End If
    Next sc
    On Error GoTo 0
End Function

Private Function IsAxisSlicerCache(ByVal sc As SlicerCache, ByVal fieldName As String) As Boolean
    Dim sl As Slicer
    Dim srcName As String

    IsAxisSlicerCache = False

    On Error Resume Next
    srcName = CStr(sc.SourceName)
    If srcName <> "" Then
        If StrComp(srcName, fieldName, vbTextCompare) = 0 Then
            IsAxisSlicerCache = True
            Exit Function
        End If
        If InStr(1, srcName, fieldName, vbTextCompare) > 0 Then
            IsAxisSlicerCache = True
            Exit Function
        End If
    End If

    For Each sl In sc.Slicers
        If StrComp(CStr(sl.Caption), fieldName, vbTextCompare) = 0 Then
            IsAxisSlicerCache = True
            Exit Function
        End If
        If InStr(1, CStr(sl.Name), fieldName, vbTextCompare) > 0 Then
            IsAxisSlicerCache = True
            Exit Function
        End If
    Next sl
    On Error GoTo 0
End Function

Private Function GetSingleSelectedAxisFromSlicer(ByVal fieldName As String) As String
    Dim sc As SlicerCache
    Dim targetSc As SlicerCache
    Dim sl As Slicer
    Dim si As SlicerItem
    Dim visList As Variant
    Dim selectedCount As Long
    Dim selectedName As String
    Dim secondName As String
    Dim preferName As String
    Dim selectedNamesLog As String

    GetSingleSelectedAxisFromSlicer = ""
    Set targetSc = Nothing

    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        For Each sl In sc.Slicers
            If StrComp(CStr(sl.Name), "Sl_" & fieldName, vbTextCompare) = 0 Then
                Set targetSc = sc
                Exit For
            End If
        Next sl
        If Not targetSc Is Nothing Then Exit For
    Next sc

    If targetSc Is Nothing Then
        For Each sc In ThisWorkbook.SlicerCaches
            If IsAxisSlicerCache(sc, fieldName) Then
                Set targetSc = sc
                Exit For
            End If
        Next sc
    End If

    If targetSc Is Nothing Then
        WriteDetailDebugLog "[Detail] slicer cache not found for field=" & fieldName
        On Error GoTo 0
        Exit Function
    End If

    ' 一部環境では SlicerItem.Selected が全件 True になるため、VisibleSlicerItemsList を優先して実選択を判定する
    If TryGetSingleVisibleSlicerItem(targetSc, GetSingleSelectedAxisFromSlicer) Then
        WriteDetailDebugLog "[Detail] slicer single(visible list priority) field=" & fieldName & " value=" & GetSingleSelectedAxisFromSlicer
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    selectedCount = 0
    selectedName = ""
    secondName = ""
    preferName = ""
    selectedNamesLog = ""
    For Each si In targetSc.SlicerItems
        If si.Selected Then
            If Not IsBlankLikeItemName(CStr(si.Name)) Then
                selectedCount = selectedCount + 1
                If selectedCount = 1 Then
                    selectedName = CStr(si.Name)
                ElseIf selectedCount = 2 Then
                    secondName = CStr(si.Name)
                End If

                If selectedNamesLog <> "" Then selectedNamesLog = selectedNamesLog & " | "
                selectedNamesLog = selectedNamesLog & CStr(si.Name)

                If m_DetailAxisField = fieldName And m_DetailAxisValue <> "" Then
                    If StrComp(CStr(si.Name), m_DetailAxisValue, vbTextCompare) <> 0 Then
                        preferName = CStr(si.Name)
                    End If
                End If
            End If
        End If
    Next si

    If Err.Number = 0 And selectedCount = 1 Then
        GetSingleSelectedAxisFromSlicer = selectedName
        WriteDetailDebugLog "[Detail] slicer single(selected flag) field=" & fieldName & " value=" & GetSingleSelectedAxisFromSlicer
        On Error GoTo 0
        Exit Function
    End If

    If Err.Number = 0 And selectedCount > 1 Then
        If selectedCount = 2 Then
            If preferName <> "" Then
                GetSingleSelectedAxisFromSlicer = preferName
            ElseIf m_DetailAxisField = fieldName And m_DetailAxisValue <> "" Then
                If StrComp(selectedName, m_DetailAxisValue, vbTextCompare) = 0 And secondName <> "" Then
                    GetSingleSelectedAxisFromSlicer = secondName
                Else
                    GetSingleSelectedAxisFromSlicer = selectedName
                End If
            Else
                GetSingleSelectedAxisFromSlicer = selectedName
            End If

            NormalizeAxisSlicerSelection targetSc, GetSingleSelectedAxisFromSlicer
            WriteDetailDebugLog "[Detail] slicer multi-selected field=" & fieldName & " count=" & selectedCount & " pick=" & GetSingleSelectedAxisFromSlicer & " names=" & selectedNamesLog
            On Error GoTo 0
            Exit Function
        Else
            WriteDetailDebugLog "[Detail] slicer multi-selected unresolved field=" & fieldName & " count=" & selectedCount & " names=" & selectedNamesLog
            On Error GoTo 0
            Exit Function
        End If
    End If

    WriteDetailDebugLog "[Detail] slicer selected-count field=" & fieldName & " count=" & selectedCount & " first=" & selectedName & " second=" & secondName

    If TryGetSingleVisibleSlicerItem(targetSc, GetSingleSelectedAxisFromSlicer) Then
        WriteDetailDebugLog "[Detail] slicer single(visible list) field=" & fieldName & " value=" & GetSingleSelectedAxisFromSlicer
    End If

    If GetSingleSelectedAxisFromSlicer = "" Then
        WriteDetailDebugLog "[Detail] slicer selection unresolved field=" & fieldName
    End If

    On Error GoTo 0
End Function

Private Function TryGetSingleVisibleSlicerItem(ByVal sc As SlicerCache, ByRef itemName As String) As Boolean
    Dim visList As Variant
    Dim i As Long
    Dim cnt As Long
    Dim oneValue As String

    itemName = ""
    TryGetSingleVisibleSlicerItem = False

    On Error Resume Next
    visList = sc.VisibleSlicerItemsList
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If Not IsArray(visList) Then Exit Function

    cnt = 0
    oneValue = ""
    For i = LBound(visList) To UBound(visList)
        oneValue = ExtractSlicerUniqueNameValue(CStr(visList(i)))
        If Not IsBlankLikeItemName(oneValue) Then
            cnt = cnt + 1
            itemName = oneValue
            If cnt > 1 Then Exit For
        End If
    Next i

    If cnt = 1 Then
        TryGetSingleVisibleSlicerItem = True
    Else
        itemName = ""
    End If
End Function

Private Sub NormalizeAxisSlicerSelection(ByVal sc As SlicerCache, ByVal keepName As String)
    Dim si As SlicerItem

    If Trim$(keepName) = "" Then Exit Sub

    On Error Resume Next
    For Each si In sc.SlicerItems
        If si.Selected Then
            If Not IsBlankLikeItemName(CStr(si.Name)) Then
                If StrComp(CStr(si.Name), keepName, vbTextCompare) <> 0 Then
                    si.Selected = False
                End If
            End If
        End If
    Next si
    On Error GoTo 0
End Sub

Private Function ExtractSlicerUniqueNameValue(ByVal uniqueName As String) As String
    Dim p1 As Long
    Dim p2 As Long

    ExtractSlicerUniqueNameValue = ""
    If Trim$(uniqueName) = "" Then Exit Function

    p2 = InStrRev(uniqueName, "]")
    p1 = InStrRev(uniqueName, "[")
    If p1 > 0 And p2 > p1 Then
        ExtractSlicerUniqueNameValue = Mid$(uniqueName, p1 + 1, p2 - p1 - 1)
    End If
End Function

Private Sub ApplyRoleReasonDetail(ByVal ws As Worksheet, ByVal mainPt As PivotTable, ByVal mainCho As ChartObject, ByVal axisField As String, ByVal axisValue As String, ByVal pNum As Integer)
    Dim detailPt As PivotTable
    Dim detailCho As ChartObject
    Dim pfAxis As PivotField
    Dim pfKubun As PivotField
    Dim pfValue As PivotField
    Dim reasonCount As Long
    Dim hasData As Boolean
    Dim valueSet As Boolean
    Dim titleAxisValue As String
    Dim oldEnableEvents As Boolean
    Dim resolvedAxisValue As String

    oldEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo SafeExit

    Set detailPt = GetOrCreateDetailPivot(ws, mainPt.PivotCache)
    If detailPt Is Nothing Then GoTo SafeExit

    detailPt.ManualUpdate = True
    detailPt.HasAutoFormat = False
    detailPt.PreserveFormatting = True
    ResetDetailPivotLayout detailPt
    ConnectDetailPivotToSlicers detailPt
    WriteDetailDebugLog "[Detail] Apply start axisField=" & axisField & " axisValue(in)=" & axisValue

    ' 2段階目の値は AddDataField で再作成する（Name代入エラー回避）。
    valueSet = ConfigureDetailDataField(detailPt, pNum)

    If Not valueSet Then
        WriteDetailDebugLog "[Detail] value field unresolved; skip update"
        GoTo SafeExit
    End If
    Dim srs As Series

    On Error Resume Next
    Set pfKubun = detailPt.PivotFields("区分")
    Set pfAxis = detailPt.PivotFields(axisField)
    On Error GoTo SafeExit
    If pfKubun Is Nothing Or pfAxis Is Nothing Then
        WriteDetailDebugLog "[Detail] pivot field unresolved kubun=" & CStr(Not pfKubun Is Nothing) & " axis=" & CStr(Not pfAxis Is Nothing)
        GoTo SafeExit
    End If
    pfKubun.Orientation = xlPageField
    pfKubun.Position = 1
    On Error Resume Next
    pfKubun.CurrentPage = IIf(P_Kubun2 = 2, "期待", "きっかけ")
    On Error GoTo 0

    pfAxis.Orientation = xlPageField
    pfAxis.Position = 2

    resolvedAxisValue = ResolveAxisValueForPageField(detailPt, axisField, axisValue)
    If Trim$(resolvedAxisValue) <> "" Then
        On Error Resume Next
        pfAxis.CurrentPage = resolvedAxisValue
        If Err.Number <> 0 Then
            Err.Clear
            axisValue = GetFirstVisiblePivotItem(detailPt, axisField)
            If axisValue <> "" Then
                pfAxis.CurrentPage = axisValue
            End If
            If Err.Number <> 0 Then
                Err.Clear
                pfAxis.ClearAllFilters
                axisValue = ""
            End If
        End If
        On Error GoTo 0
    Else
        axisValue = ""
    End If

    On Error Resume Next
    WriteDetailDebugLog "[Detail] Apply axisField=" & axisField & " axisValue(in)=" & axisValue & " pfAxis.CurrentPage=" & CStr(pfAxis.CurrentPage)
    On Error GoTo 0

    On Error Resume Next
    detailPt.PivotFields("集計理由項目").Orientation = xlRowField
    If Err.Number <> 0 Then
        WriteDetailDebugLog "[Detail] set reason row failed=" & CStr(Err.Number) & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo SafeExit

    detailPt.ManualUpdate = False
    detailPt.RefreshTable
    WriteDetailDebugLog "[Detail] refresh done"

    On Error Resume Next
    hasData = Not detailPt.DataBodyRange Is Nothing
    On Error GoTo 0
    If Not hasData Then
        detailPt.ManualUpdate = True
        pfAxis.ClearAllFilters
        detailPt.ManualUpdate = False
        detailPt.RefreshTable
        On Error Resume Next
        hasData = Not detailPt.DataBodyRange Is Nothing
        On Error GoTo 0
        If Not hasData Then
            WriteDetailDebugLog "[Detail] hasData=false after retry; keep previous chart"
            GoTo SafeExit
        End If
    End If

    reasonCount = CountVisiblePivotItems(detailPt, "集計理由項目")
    If reasonCount = 0 Then
        ' 単一軸で0件の場合は軸フィルタを解除して再取得
        detailPt.ManualUpdate = True
        pfAxis.ClearAllFilters
        detailPt.ManualUpdate = False
        detailPt.RefreshTable
        reasonCount = CountVisiblePivotItems(detailPt, "集計理由項目")
        If reasonCount = 0 Then
            WriteDetailDebugLog "[Detail] reasonCount=0 after retry; keep previous chart"
            GoTo SafeExit
        End If
    End If

    m_DetailAxisField = axisField
    m_DetailAxisValue = axisValue

    titleAxisValue = ""
    On Error Resume Next
    titleAxisValue = CStr(pfAxis.CurrentPage)
    If titleAxisValue = "(All)" Then titleAxisValue = ""
    On Error GoTo 0

    If titleAxisValue = "" Then
        titleAxisValue = Trim$(axisValue)
    End If
    If titleAxisValue = "" Then
        titleAxisValue = GetSingleSelectedAxisFromSlicer(axisField)
    End If
    If titleAxisValue = "" Then
        titleAxisValue = GetSingleSelectedRole(detailPt, axisField)
    End If
    If titleAxisValue = "" Then
        titleAxisValue = "全体"
    End If
    WriteDetailDebugLog "[Detail] titleAxisValue(final)=" & titleAxisValue
    If titleAxisValue <> "全体" Then
        m_DetailAxisValue = titleAxisValue
    End If

    Set detailCho = GetOrCreateDetailChart(ws, mainCho)
    If detailCho Is Nothing Then
        WriteDetailDebugLog "[Detail] detail chart create failed"
        GoTo SafeExit
    End If

    With detailCho.Chart
        .SetSourceData Source:=detailPt.TableRange1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = fnTrimCode(P_Pattern) & "（" & titleAxisValue & " / 理由詳細）"

        On Error Resume Next
        .ApplyDataLabels Type:=xlDataLabelsShowValue
        For Each srs In .SeriesCollection
            srs.HasDataLabels = True
            srs.DataLabels.ShowValue = True
        Next srs

        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        On Error GoTo 0
    End With
    WriteDetailDebugLog "[Detail] chart updated"

SafeExit:
    If Err.Number <> 0 Then
        WriteDetailDebugLog "[Detail] Apply error=" & CStr(Err.Number) & " " & Err.Description
        Err.Clear
    End If
    Application.EnableEvents = oldEnableEvents
End Sub

Private Function ConfigureDetailDataField(ByVal pt As PivotTable, ByVal pNum As Integer) As Boolean
    Dim srcField As PivotField
    Dim dataField As PivotField
    Dim desiredSource As String
    Dim desiredCaption As String
    Dim desiredFunc As XlConsolidationFunction

    ConfigureDetailDataField = False

    If pNum = 3 Or pNum = 4 Then
        desiredSource = "社員番号"
        desiredCaption = "理由延べ数 "
        desiredFunc = xlCount
    Else
        desiredSource = "人数フラグ"
        desiredCaption = "退職者数 "
        desiredFunc = xlSum
    End If

    ClearDetailDataFields pt

    On Error Resume Next
    Set srcField = pt.PivotFields(desiredSource)
    If Not srcField Is Nothing Then
        Set dataField = pt.AddDataField(srcField, desiredCaption, desiredFunc)
        If Not dataField Is Nothing Then
            dataField.NumberFormat = "#,##0"
            ConfigureDetailDataField = True
        End If
    End If

    If Not ConfigureDetailDataField Then
        Set srcField = Nothing
        Set srcField = pt.PivotFields("人数フラグ")
        If srcField Is Nothing Then Set srcField = pt.PivotFields("社員番号")
        If Not srcField Is Nothing Then
            If StrComp(CStr(srcField.Name), "社員番号", vbTextCompare) = 0 Then
                Set dataField = pt.AddDataField(srcField, "理由延べ数 ", xlCount)
            Else
                Set dataField = pt.AddDataField(srcField, "退職者数 ", xlSum)
            End If
            If Not dataField Is Nothing Then
                dataField.NumberFormat = "#,##0"
                ConfigureDetailDataField = True
            End If
        End If
    End If

    If Not ConfigureDetailDataField Then
        WriteDetailDebugLog "[Detail] ConfigureDataField failed source=" & desiredSource
    End If
    On Error GoTo 0
End Function

Private Sub ClearDetailDataFields(ByVal pt As PivotTable)
    Dim df As PivotField

    On Error Resume Next
    For Each df In pt.DataFields
        df.Orientation = xlHidden
    Next df
    On Error GoTo 0
End Sub

Private Function ResolveAxisValueForPageField(ByVal pt As PivotTable, ByVal fieldName As String, ByVal axisValue As String) As String
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim src As String
    Dim srcTrim As String

    ResolveAxisValueForPageField = ""
    src = Trim$(axisValue)
    If src = "" Then Exit Function

    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0
    If pf Is Nothing Then Exit Function

    On Error Resume Next
    Set pi = pf.PivotItems(src)
    If Not pi Is Nothing Then
        ResolveAxisValueForPageField = src
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    srcTrim = fnTrimCode(src)
    If srcTrim = "" Then Exit Function

    On Error Resume Next
    For Each pi In pf.PivotItems
        If StrComp(fnTrimCode(CStr(pi.Name)), srcTrim, vbTextCompare) = 0 Then
            ResolveAxisValueForPageField = CStr(pi.Name)
            Exit For
        End If
    Next pi
    On Error GoTo 0
End Function

Private Sub ResetDetailDebugLog()
    If Not P_DebugLogEnabled Then Exit Sub

    On Error Resume Next
    m_DebugRow = 3
    With ThisWorkbook.Worksheets("分析グラフ")
        .Range("Y2").Value = "DetailDebug"
        .Range("Y3:Y80").ClearContents
    End With
    On Error GoTo 0
End Sub

Private Sub WriteDetailDebugLog(ByVal msg As String)
    Debug.Print msg
    If Not P_DebugLogEnabled Then Exit Sub

    On Error Resume Next
    If m_DebugRow < 3 Then m_DebugRow = 3
    ThisWorkbook.Worksheets("分析グラフ").Range("Y" & CStr(m_DebugRow)).Value = msg
    m_DebugRow = m_DebugRow + 1
    On Error GoTo 0
End Sub

Private Function GetAxisTitleFromSlicer(ByVal fieldName As String) As String
    Dim sc As SlicerCache
    Dim si As SlicerItem
    Dim selectedCount As Long
    Dim totalCount As Long
    Dim selectedName As String

    GetAxisTitleFromSlicer = ""

    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        If IsAxisSlicerCache(sc, fieldName) Then
            selectedCount = 0
            totalCount = 0
            selectedName = ""

            For Each si In sc.SlicerItems
                totalCount = totalCount + 1
                If si.Selected Then
                    selectedCount = selectedCount + 1
                    selectedName = CStr(si.Name)
                End If
            Next si

            If selectedCount = 1 Then
                GetAxisTitleFromSlicer = selectedName
            ElseIf selectedCount = 0 Or selectedCount = totalCount Then
                GetAxisTitleFromSlicer = "全体"
            Else
                GetAxisTitleFromSlicer = "複数選択"
            End If

            On Error GoTo 0
            Exit Function
        End If
    Next sc
    On Error GoTo 0
End Function

Private Function GetSingleSelectedRole(ByVal pt As PivotTable, ByVal fieldName As String) As String
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim cnt As Long
    Dim selectedName As String

    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0

    If pf Is Nothing Then
        GetSingleSelectedRole = ""
        Exit Function
    End If

    If pf.Orientation = xlPageField Then
        On Error Resume Next
        selectedName = CStr(pf.CurrentPage)
        On Error GoTo 0
        If selectedName <> "" And selectedName <> "(All)" Then
            GetSingleSelectedRole = selectedName
        Else
            cnt = 0
            On Error Resume Next
            For Each pi In pf.PivotItems
                If pi.Visible Then
                    If Not IsBlankLikeItemName(CStr(pi.Name)) Then
                        selectedName = CStr(pi.Name)
                        cnt = cnt + 1
                        If cnt > 1 Then Exit For
                    End If
                End If
            Next pi
            On Error GoTo 0

            If cnt = 1 Then
                GetSingleSelectedRole = selectedName
            Else
                GetSingleSelectedRole = ""
            End If
        End If
        Exit Function
    End If

    cnt = 0
    For Each pi In pf.PivotItems
        If pi.Visible Then
            If Not IsBlankLikeItemName(CStr(pi.Name)) Then
                selectedName = CStr(pi.Name)
                cnt = cnt + 1
                If cnt > 1 Then Exit For
            End If
        End If
    Next pi

    If cnt = 1 Then
        GetSingleSelectedRole = selectedName
    Else
        GetSingleSelectedRole = ""
    End If
End Function

Private Function IsBlankLikeItemName(ByVal itemName As String) As Boolean
    Dim n As String
    Dim s As String

    n = Trim$(itemName)
    s = LCase$(n)
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "（", "")
    s = Replace(s, "）", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(12288), "")

    IsBlankLikeItemName = (n = "" Or InStr(s, "blank") > 0 Or InStr(s, "空白") > 0)
End Function

Private Function ShouldUseTwoStepReason(ByVal pt As PivotTable, ByVal pNum As Integer) As Boolean
    If pNum < 3 Or pNum > 8 Then
        ShouldUseTwoStepReason = False
        Exit Function
    End If

    ShouldUseTwoStepReason = (CountVisiblePivotItems(pt, "集計理由項目") > MAX_REASON_ITEMS)
End Function

Private Function GetAxisFieldByPattern(ByVal pNum As Integer) As String
    Select Case pNum
        Case 1, 3: GetAxisFieldByPattern = "性別名"
        Case 2, 4: GetAxisFieldByPattern = "勤続区分"
        Case 5: GetAxisFieldByPattern = "役職名"
        Case 6: GetAxisFieldByPattern = "会社名"
        Case 7: GetAxisFieldByPattern = "工場名"
        Case 8: GetAxisFieldByPattern = "所属名"
        Case 9, 11: GetAxisFieldByPattern = "退職月"
        Case 10, 12: GetAxisFieldByPattern = "退職年"
        Case Else: GetAxisFieldByPattern = ""
    End Select
End Function

Private Sub ApplyTopNToReason(ByVal pt As PivotTable, ByVal topN As Long)
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields("集計理由項目")
    On Error GoTo 0

    If pf Is Nothing Then Exit Sub

    On Error Resume Next
    pf.ClearAllFilters
    pf.AutoSort xlDescending, pt.DataFields(1).name
    pf.PivotFilters.Add2 Type:=xlTopCount, DataField:=pt.DataFields(1), Value1:=topN
    If Err.Number <> 0 Then
        Err.Clear
        pf.PivotFilters.Add Type:=xlTopCount, DataField:=pt.DataFields(1), Value1:=topN
    End If
    On Error GoTo 0
End Sub

Private Function CountVisiblePivotItems(ByVal pt As PivotTable, ByVal fieldName As String) As Long
    Dim pf As PivotField
    Dim pi As PivotItem

    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0

    If pf Is Nothing Then
        CountVisiblePivotItems = 0
        Exit Function
    End If

    CountVisiblePivotItems = 0
    On Error Resume Next
    For Each pi In pf.PivotItems
        If pi.Visible Then
            CountVisiblePivotItems = CountVisiblePivotItems + 1
        End If
    Next pi
    On Error GoTo 0
End Function

Private Function GetOrCreateDetailPivot(ByVal ws As Worksheet, ByVal pcCache As PivotCache) As PivotTable
    Dim pt As PivotTable

    On Error Resume Next
    Set pt = ws.PivotTables(DETAIL_PIVOT_NAME)
    On Error GoTo 0

    If pt Is Nothing Then
        On Error Resume Next
        Set pt = pcCache.CreatePivotTable(TableDestination:=ws.Range("BA100"), TableName:=DETAIL_PIVOT_NAME)
        On Error GoTo 0
    End If

    Set GetOrCreateDetailPivot = pt
End Function

Private Function GetOrCreateDetailChart(ByVal ws As Worksheet, ByVal mainCho As ChartObject) As ChartObject
    Dim cho As ChartObject

    On Error Resume Next
    Set cho = ws.ChartObjects(DETAIL_CHART_NAME)
    On Error GoTo 0

    If cho Is Nothing Then
        Set cho = ws.ChartObjects.Add(Left:=mainCho.Left, Top:=mainCho.Top + mainCho.Height + 20, Width:=mainCho.Width, Height:=360)
        cho.name = DETAIL_CHART_NAME
    Else
        cho.Left = mainCho.Left
        cho.Top = mainCho.Top + mainCho.Height + 20
        cho.Width = mainCho.Width
        cho.Height = 360
    End If

    Set GetOrCreateDetailChart = cho
End Function

Private Sub ResetDetailPivotLayout(ByVal pt As PivotTable)
    On Error Resume Next
    ' スライサー選択を保持するため、ここでは全体フィルタ解除を行わない
    HidePivotFieldIfExists pt, "人数フラグ"
    HidePivotFieldIfExists pt, "社員番号"
    HidePivotFieldIfExists pt, "区分"
    HidePivotFieldIfExists pt, "集計理由項目"
    HidePivotFieldIfExists pt, "性別名"
    HidePivotFieldIfExists pt, "勤続区分"
    HidePivotFieldIfExists pt, "役職名"
    HidePivotFieldIfExists pt, "会社名"
    HidePivotFieldIfExists pt, "工場名"
    HidePivotFieldIfExists pt, "所属名"
    HidePivotFieldIfExists pt, "退職月"
    HidePivotFieldIfExists pt, "退職年"
    On Error GoTo 0
End Sub

Private Sub HidePivotFieldIfExists(ByVal pt As PivotTable, ByVal fieldName As String)
    On Error Resume Next
    pt.PivotFields(fieldName).Orientation = xlHidden
    On Error GoTo 0
End Sub

Private Sub ConnectDetailPivotToSlicers(ByVal pt As PivotTable)
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.PivotTables.AddPivotTable pt
    Next sc
    On Error GoTo 0
End Sub

Private Sub RebindSlicersToDetailPivot(ByVal mainPt As PivotTable, ByVal detailPt As PivotTable)
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.PivotTables.AddPivotTable detailPt
        sc.PivotTables.RemovePivotTable mainPt
    Next sc
    On Error GoTo 0
End Sub

Private Sub RemoveDetailArtifacts(ByVal ws As Worksheet)
    On Error Resume Next
    ws.ChartObjects(DETAIL_CHART_NAME).Delete
    On Error GoTo 0
End Sub

Private Function DetailChartExists(ByVal ws As Worksheet) As Boolean
    Dim cho As ChartObject

    On Error Resume Next
    Set cho = ws.ChartObjects(DETAIL_CHART_NAME)
    On Error GoTo 0

    DetailChartExists = Not cho Is Nothing
End Function

Private Function IsPivotItemVisible(ByVal pt As PivotTable, ByVal fieldName As String, ByVal itemName As String) As Boolean
    Dim pf As PivotField
    Dim pi As PivotItem

    IsPivotItemVisible = False
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    If pf Is Nothing Then Exit Function

    Set pi = pf.PivotItems(itemName)
    If pi Is Nothing Then Exit Function

    IsPivotItemVisible = pi.Visible
    On Error GoTo 0
End Function

Private Function GetFirstVisiblePivotItem(ByVal pt As PivotTable, ByVal fieldName As String) As String
    Dim pf As PivotField
    Dim pi As PivotItem

    GetFirstVisiblePivotItem = ""
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    If pf Is Nothing Then Exit Function

    For Each pi In pf.PivotItems
        If pi.Visible Then
            GetFirstVisiblePivotItem = CStr(pi.name)
            Exit For
        End If
    Next pi
    On Error GoTo 0
End Function



