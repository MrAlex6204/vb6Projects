VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMDATEST2 
   BackColor       =   &H00404040&
   Caption         =   "RESUMEN x TIPO DE CUENTA"
   ClientHeight    =   7185
   ClientLeft      =   -1575
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   735
      Left            =   120
      Picture         =   "FRMDATEST2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "RESUMEN:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   5655
      Begin VB.CommandButton CMDRES 
         Caption         =   "&RESUMEN"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox TXTING 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXTCLI 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL DE INGRESOS POR CUENTAS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL DE CLIENTES CON CUENTA EN EL BANCO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4005
      End
   End
   Begin VB.CommandButton CMDING 
      Caption         =   "&INGRESOS x TIPO_CUENTA"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton CMDCOUNT 
      Caption         =   "&CLIENTES x TIPO_CUENTA"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFDATA2 
      Height          =   2415
      Left            =   5160
      TabIndex        =   1
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFDATA 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   4210752
      ForeColorFixed  =   16777215
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
Attribute VB_Name = "FRMDATEST2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDCOUNT_Click()
If RS.State = 1 Then RS.Close
SQL = "SELECT TC.DESTCTA AS [TIPO DE CUENTA],COUNT(*) AS [TOTAL DE CLIENTES] FROM TIPOCUENTA TC,CUENTA C WHERE TC.CODTCTA=C.CODTCTA GROUP BY TC.DESTCTA"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFDATA.DataSource = RS
MHFDATA.ColWidth(0) = 1800
MHFDATA.ColWidth(1) = 2000
MHFDATA.ColAlignment = 3
End Sub

Private Sub CMDING_Click()
If RS.State = 1 Then RS.Close
SQL = "SELECT TC.DESTCTA AS [TIPO DE CUENTA],SUM(MONTO) AS [TOTAL DE INGRESOS] FROM TIPOCUENTA TC,CUENTA C WHERE TC.CODTCTA=C.CODTCTA GROUP BY TC.DESTCTA"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFDATA2.DataSource = RS
MHFDATA2.ColWidth(0) = 1800
MHFDATA2.ColWidth(1) = 2100
MHFDATA2.ColAlignment = 3
End Sub

Private Sub CMDRES_Click()
TXTCLI = SUMACOLUMNAMHF(1, MHFDATA)
TXTING = SUMACOLUMNAMHF(1, MHFDATA2)
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
End Sub

