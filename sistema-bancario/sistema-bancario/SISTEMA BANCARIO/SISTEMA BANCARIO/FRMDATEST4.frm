VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMDATEST4 
   BackColor       =   &H00404040&
   Caption         =   "CUENTAS x MESES"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   9300
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   735
      Left            =   120
      Picture         =   "FRMDATEST4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFRC 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(*) ESTE CUADRO DE REFERENCIA CRUZADA MUESTRA LA CANTIDAD DE TIPOS DE CUENTA CREADAS POR CADA MES DEL AÑO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   7695
   End
End
Attribute VB_Name = "FRMDATEST4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "TRANSFORM COUNT(FECHAREG) SELECT T.DESTCTA AS [TIPO_CUENTA] FROM TIPOCUENTA T,CUENTA C WHERE T.CODTCTA=C.CODTCTA GROUP BY T.DESTCTA PIVOT FORMAT([FECHAREG],'MMM') IN ('ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SET','OCT','NOV','DIC') "
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFRC.DataSource = RS
MHFRC.ColWidth(0) = 1800
End Sub
