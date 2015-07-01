VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMDATEST3 
   BackColor       =   &H00404040&
   Caption         =   "RESUMEN DE CLIENTES x DISTRITOS"
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDINGDIS 
      Caption         =   "&INGRESOS x DISTRITO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton CMDCC 
      Caption         =   "&CLIENTES x DISTRITO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFING 
      Height          =   4575
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8070
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
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   735
      Left            =   120
      Picture         =   "FRMDATEST3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFDIS 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8070
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
      Caption         =   "(*) EN LA TABLA INGRESOS, SOLO FIGURAN LOS DISTRITOS QUE POSEEN CLIENTES CON CUENTAS."
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
      TabIndex        =   5
      Top             =   5520
      Width           =   6735
   End
End
Attribute VB_Name = "FRMDATEST3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCC_Click()
If RS.State = 1 Then RS.Close
SQL = "SELECT D.DISTRITO AS [DISTRITO], COUNT(*) AS [TOTAL DE CLIENTES] FROM CLIENTES C, DISTRITOS D WHERE C.CODDIS=D.CODDIS GROUP BY D.DISTRITO"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFDIS.DataSource = RS
MHFDIS.ColAlignment = 3
MHFDIS.ColAlignmentFixed = 3
MHFDIS.ColWidth(0) = 2300
MHFDIS.ColWidth(1) = 2000
End Sub

Private Sub CMDINGDIS_Click()
If RS.State = 1 Then RS.Close
SQL = "SELECT D.DISTRITO AS [DISTRITO], SUM(MONTO) AS [INGRESOS] FROM CLIENTES C, DISTRITOS D, CUENTA T WHERE C.CODDIS=D.CODDIS AND T.CODCLI=C.CODCLI GROUP BY D.DISTRITO"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFING.DataSource = RS
MHFING.ColAlignment = 3
MHFING.ColAlignmentFixed = 3
MHFING.ColWidth(0) = 2300
MHFING.ColWidth(1) = 2000
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
End Sub
