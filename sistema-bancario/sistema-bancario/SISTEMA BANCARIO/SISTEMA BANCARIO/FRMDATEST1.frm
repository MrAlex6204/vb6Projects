VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMDATEST1 
   BackColor       =   &H00404040&
   Caption         =   "MOVIMIENTOS REALIZADOS POR MES"
   ClientHeight    =   8490
   ClientLeft      =   -1575
   ClientTop       =   555
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "&SALIR"
      Height          =   735
      Left            =   240
      Picture         =   "FRMDATEST1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1335
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton CMDMES 
         Caption         =   "&ENERO"
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
         Index           =   0
         Left            =   240
         MaskColor       =   &H0000FF00&
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&FEBRERO"
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
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&MARZO"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&ABRIL"
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
         Index           =   3
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "MA&YO"
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
         Index           =   4
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&JUNIO"
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
         Index           =   5
         Left            =   2880
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "JU&LIO"
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
         Index           =   6
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "A&GOSTO"
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
         Index           =   7
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&SETIEMBRE"
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
         Index           =   8
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&OCTUBRE"
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
         Index           =   9
         Left            =   5520
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&NOVIEMBRE"
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
         Index           =   10
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDMES 
         Caption         =   "&DICIEMBRE"
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
         Index           =   11
         Left            =   6840
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFMES 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7858
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
      Caption         =   "(*) PARA VISUALIZAR LOS MOVIMIENTOS EN UN DETERMINADO MES, SELECCIONE UNO DE LOS BOTONES DE COMANDO."
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
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   5895
   End
End
Attribute VB_Name = "FRMDATEST1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDMES_Click(Index As Integer)
Dim SQL As String
If RS.State = 1 Then RS.Close
Select Case Index
       Case 0:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('ENE')"
       Case 1:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('FEB')"
       Case 2:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('MAR')"
       Case 3:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('ABR')"
       Case 4:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('MAY')"
       Case 5:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('JUN')"
       Case 6:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('JUL')"
       Case 7:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('AGO')"
       Case 8:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('SEP')"
       Case 9:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('OCT')"
       Case 10:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('NOV')"
       Case 11:
       SQL = "SELECT NROCTA AS [NRO_CTA],O.DESOPE AS [OPERACION], MONTOMOV AS [MONTO], FECHAMOV AS [FECHA],HORAMOV AS [HORA] FROM OPBANCARIA O, MOVIMIENTOS M WHERE M.CODOPE=O.CODOPE AND FORMAT([FECHAMOV],'mmm') IN ('DIC')"
End Select
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFMES.DataSource = RS
End Sub

Private Sub CMDSALIR_Click()
Unload Me
MDISIS.Show
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
SQL = "SELECT*FROM MOVIMIENTOS"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
MHFMES.ColWidth(1) = 1200
End Sub
