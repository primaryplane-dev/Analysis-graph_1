Option Explicit

' --- カレンダーフォーム用の共有変数 ---
Public P_DATEC As Date           ' 選択された日付
Public P_Calendar_FLG As Boolean ' 決定ボタンが押されたかどうか

' --- 抽出・分析用の共有変数 ---
Public P_DateFrom As String      ' 抽出開始日 (YYYYMMDD)
Public P_DateTo As String        ' 抽出終了日 (YYYYMMDD)
Public P_Kubun2 As Integer       ' 1:きっかけ, 2:期待
Public P_Pattern As String       ' コンボボックスで選んだ13パターンの名前
Public P_Regist As Boolean       ' フォームで実行が押されたかどうかのフラグ

' --- 追跡用辞書オブジェクトの定義 ---
Public dicKaisyaName As Object
Public dicKojoName As Object
Public dicPosName As Object
Public dicDeptName As Object
