VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMCON1 
   BackColor       =   &H00404040&
   Caption         =   "CONSULTA POR CLIENTES"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10245
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   1095
      Left            =   1200
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CMDIMPRIMIR 
      Caption         =   "&IMPRIMIR REPORTE"
      Height          =   1095
      Left            =   120
      Picture         =   "Form3.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFCLI 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
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
Attribute VB_Name = "FRMCON1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDIMPRIMIR_Click()
Set DRPCLIENTES.DataSource = RS
DRPCLIENTES.Show
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "SELECT CODCLI AS [CODIGO], NOMCLI AS [NOMBRE], DIRCLI AS [DIRECCION], TELCLI AS [TELEFONO], SEXO, D.DISTRITO FROM CLIENTES C, DISTRITOS D WHERE C.CODDIS=D.CODDIS"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFCLI.DataSource = RS
MHFCLI.ColWidth(1) = 2400
MHFCLI.ColWidth(2) = 4300
MHFCLI.ColWidth(3) = 1200
MHFCLI.ColWidth(5) = 2400
End Sub

