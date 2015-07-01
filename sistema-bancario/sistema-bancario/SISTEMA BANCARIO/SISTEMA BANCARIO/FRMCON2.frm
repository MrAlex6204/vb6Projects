VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMCON2 
   BackColor       =   &H00404040&
   Caption         =   "CONSULTA POR EMPLEADOS"
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
      Picture         =   "FRMCON2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   1095
      Left            =   1200
      Picture         =   "FRMCON2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFEMP 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
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
Attribute VB_Name = "FRMCON2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDIMPRIMIR_Click()
Set DRPEMP.DataSource = RS
DRPEMP.Show
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "SELECT CODEMP AS [CODIGO],EMPLEADO,C.DESCAR AS [CARGO],DIR AS [DIRECCION],TEL AS [TELEFONO],SEXO,D.DISTRITO FROM EMPLEADOS E, CARGOS C, DISTRITOS D WHERE E.CODDIS=D.CODDIS AND E.CODCAR=C.CODCAR"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFEMP.DataSource = RS
MHFEMP.ColWidth(1) = 2000
MHFEMP.ColWidth(2) = 2000
MHFEMP.ColWidth(3) = 4300
MHFEMP.ColWidth(4) = 1200
MHFEMP.ColWidth(6) = 2400
End Sub
