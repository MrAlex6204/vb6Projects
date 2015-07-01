VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMCON3 
   BackColor       =   &H00404040&
   Caption         =   "CONSULTA POR CUENTAS BANCARIAS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   1095
      Left            =   1200
      Picture         =   "FRMCON3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton CMDIMPRIMIR 
      Caption         =   "&IMPRIMIR REPORTE"
      Height          =   1095
      Left            =   120
      Picture         =   "FRMCON3.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFCTA 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   -2147483634
      BackColorBkg    =   4210752
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FRMCON3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDIMPRIMIR_Click()
Set DRPCTA.DataSource = RS
DRPCTA.Show
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "SELECT NROCTA AS [CUENTA], C.NOMCLI AS [CLIENTE], E.EMPLEADO AS [EMPLEADO], M.DESMON AS [MONEDA], T.DESTCTA AS [TIPO_CUENTA],MONTO, FECHAREG AS [FECHA_REGISTRO]FROM CUENTA CT, CLIENTES C, EMPLEADOS E, MONEDA M, TIPOCUENTA T WHERE C.CODCLI=CT.CODCLI AND E.CODEMP=CT.CODEMP AND CT.CODMON=M.CODMON AND T.CODTCTA=CT.CODTCTA"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFCTA.DataSource = RS
MHFCTA.ColWidth(0) = 1300
MHFCTA.ColWidth(1) = 2300
MHFCTA.ColWidth(2) = 2200
MHFCTA.ColWidth(4) = 1900
MHFCTA.ColWidth(2) = 2200
MHFCTA.ColWidth(5) = 800
MHFCTA.ColWidth(6) = 1800
End Sub
