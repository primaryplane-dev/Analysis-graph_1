Option Explicit

Public Const P_ConnectString As String = "Provider=IBMDA400;Data Source=HONSHA;User ID=SYSTEM;Password=FJPN2480"
'Public Const P_ConnectString As String = "Provider=IBMDA400;Data Source=FUJIPAN;User ID=SYSTEM;Password=FJPN2480"

'ini
Public P_NYNO           As Long
Public P_Kengen         As Long

'frmSelect
Public P_Regist         As Boolean
Public P_bDate          As Boolean
Public P_DateF          As Date
Public P_DateT          As Date
Public P_Kubun          As Long
Public P_Kubun2         As Long
Public P_Kubun3         As Long
Public P_SelKAI         As Long
Public P_SelKJN         As Long
Public P_Type           As Long

Public dicKaisyaName    As Dictionary
Public dicKojoName      As Dictionary
