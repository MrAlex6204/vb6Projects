VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMCON4 
   BackColor       =   &H00404040&
   Caption         =   "CONSULTA DE MOVIMIENTOS REALIZADOS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDIMPRIMIR 
      Caption         =   "&IMPRIMIR REPORTE"
      Height          =   1095
      Left            =   120
      Picture         =   "FRMCON4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   1095
      Left            =   1200
      Picture         =   "FRMCON4.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFMOV 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
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
Attribute VB_Name = "FRMCON4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDIMPRIMIR_Click()
Set DRPMOV.DataSource = RS
DRPMOV.Show
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFMOV.DataSource = RS
MHFMOV.ColWidth(1) = 1200
End Sub
