VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FRMBUS 
   BackColor       =   &H00404040&
   Caption         =   "BUSQUEDA DE DATOS"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   10905
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDSALIR 
      Height          =   615
      Left            =   6240
      Picture         =   "FRMBUS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "SALIR"
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox CBODIS 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "BUSQUEDA POR:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox CBOMON 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.ComboBox CBOOB 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox CBOTC 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPERACION BANCARIA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE CUENTA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRITOS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFBUS 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5106
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
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   8040
      Picture         =   "FRMBUS.frx":0442
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "FRMBUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBODIS_Click()
CBOMON = ""
CBOOB = ""
CBOTC = ""
SQL = "SELECT CODCLI,NOMCLI,DIRCLI,TELCLI,SEXO FROM CLIENTES C,DISTRITOS D WHERE D.CODDIS=C.CODDIS AND D.DISTRITO='" + CBODIS + "'"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFBUS.DataSource = RS
RS.Close
MHFBUS.ColWidth(0) = 1200
MHFBUS.ColWidth(1) = 2200
MHFBUS.ColWidth(2) = 4200
End Sub

Private Sub CBOMON_Click()
CBOOB = ""
CBOTC = ""
CBODIS = ""
SQL = "SELECT NROCTA,C.NOMCLI,E.EMPLEADO,T.DESTCTA,MONTO,FECHAREG FROM CLIENTES C, EMPLEADOS E,TIPOCUENTA T, CUENTA TC, MONEDA M WHERE C.CODCLI=TC.CODCLI AND TC.CODEMP=E.CODEMP AND TC.CODTCTA=T.CODTCTA AND TC.CODMON=M.CODMON AND M.DESMON='" + CBOMON + "'"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFBUS.DataSource = RS
RS.Close
MHFBUS.ColWidth(1) = 2200
MHFBUS.ColWidth(2) = 2200
MHFBUS.ColWidth(3) = 1800
MHFBUS.ColWidth(5) = 1200
End Sub

Private Sub CBOOB_Click()
CBODIS = ""
CBOMON = ""
CBOTC = ""
SQL = "SELECT NROCTA,MONTOMOV,FECHAMOV,HORAMOV FROM MOVIMIENTOS M,OPBANCARIA O WHERE O.CODOPE=M.CODOPE AND O.DESOPE='" + CBOOB + "'"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFBUS.DataSource = RS
RS.Close
MHFBUS.ColWidth(1) = 1200
MHFBUS.ColWidth(2) = 1200
MHFBUS.ColWidth(3) = 1200
End Sub

Private Sub CBOTC_Click()
CBOMON = ""
CBOOB = ""
CBODIS = ""
SQL = "SELECT NROCTA,C.NOMCLI,E.EMPLEADO,M.DESMON,MONTO,FECHAREG FROM CUENTA CT, CLIENTES C, EMPLEADOS E, MONEDA M, TIPOCUENTA T WHERE CT.CODCLI=C.CODCLI AND CT.CODEMP=E.CODEMP AND M.CODMON=CT.CODMON AND CT.CODTCTA=T.CODTCTA AND T.DESTCTA='" + CBOTC + "'"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Set MHFBUS.DataSource = RS
RS.Close
MHFBUS.ColWidth(0) = 1200
MHFBUS.ColWidth(1) = 2200
MHFBUS.ColWidth(2) = 2200
End Sub

Private Sub CMDSALIR_Click()
If MsgBox("¿ESTA SEGURO?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
   Unload Me
   MDISIS.Show
End If
End Sub

Private Sub Form_Activate()
Call COMBO("DISTRITOS", CBODIS, 1)
Call COMBO("OPBANCARIA", CBOOB, 1)
Call COMBO("TIPOCUENTA", CBOTC, 1)
Call COMBO("MONEDA", CBOMON, 1)
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
End Sub
