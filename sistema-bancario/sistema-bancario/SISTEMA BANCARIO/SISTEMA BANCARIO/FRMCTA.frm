VERSION 5.00
Begin VB.Form FRMCTA 
   BackColor       =   &H00800000&
   Caption         =   "CUENTAS REGISTRADAS"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   Icon            =   "FRMCTA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   7245
   WindowState     =   2  'Maximized
   Begin VB.Frame FRABOTONES 
      BackColor       =   &H00800000&
      Height          =   855
      Left            =   1080
      TabIndex        =   23
      Top             =   3840
      Width           =   5055
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FRMCTA.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   1
         Left            =   1320
         Picture         =   "FRMCTA.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "FRMCTA.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   3
         Left            =   3720
         Picture         =   "FRMCTA.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   6975
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "FRMCTA.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "NUEVO"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   1
         Left            =   1080
         Picture         =   "FRMCTA.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "GRABAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   2
         Left            =   2040
         Picture         =   "FRMCTA.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "BORRAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "FRMCTA.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "MODIFICAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   4
         Left            =   4920
         Picture         =   "FRMCTA.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "CANCELAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   6
         Left            =   5900
         Picture         =   "FRMCTA.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "SALIR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   5
         Left            =   3960
         Picture         =   "FRMCTA.frx":2ED6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "ACTUALIZAR"
         Top             =   230
         Width           =   975
      End
   End
   Begin VB.Frame FRADATOS 
      BackColor       =   &H00800000&
      Enabled         =   0   'False
      Height          =   3255
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   4095
      Begin VB.ComboBox CBOEMP 
         Height          =   315
         Left            =   1320
         TabIndex        =   30
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox CBOCLI 
         Height          =   315
         Left            =   1320
         TabIndex        =   28
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TXTFREG 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox TXTMON 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox CBOTC 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00800000&
         Caption         =   "TIPO DE MONEDA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton OPTMON 
            BackColor       =   &H00800000&
            Caption         =   "EUROS"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OPTMON 
            BackColor       =   &H00800000&
            Caption         =   "DOLARES"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OPTMON 
            BackColor       =   &H00800000&
            Caption         =   "SOLES"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox TXTNCTA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLEADO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   970
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA REGISTRADA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO. CUENTA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   120
      Picture         =   "FRMCTA.frx":3318
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTAS  REGISTRADAS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4350
   End
End
Attribute VB_Name = "FRMCTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCONTROL_Click(Index As Integer)
Select Case Index
       Case 0: Call LIMPIAR(Me)
               RS.MoveLast
               Call GENERACODIGO
               FRADATOS.Enabled = True
               FRABOTONES.Enabled = False
               TXTNCTA.Enabled = False
               Call HABILITARCONTROLES(False)
       Case 1: Call GRABAR
               FRADATOS.Enabled = False
               FRABOTONES.Enabled = True
               RS.Requery
               Call HABILITARCONTROLES(True)
               CMDCONTROL(5).Enabled = False
       Case 2:
       Dim CODCLI$, CODEMP$, CODTC$, MONEDA$
       If MsgBox("¿DESEA BORRAR EL SGTE. REGISTRO?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
       CODCLI = BUSCAR("CLIENTES", "NOMCLI", CBOCLI, 0)
       CODEMP = BUSCAR("EMPLEADOS", "EMPLEADO", CBOEMP, 0)
       CODTC = BUSCAR("TIPOCUENTA", "DESTCTA", CBOTC, 0)
       If OPTMON(0).Value = True Then MONEDA = "M-1"
       If OPTMON(1).Value = True Then MONEDA = "M-2"
       If OPTMON(2).Value = True Then MONEDA = "M-3"
       SQLBORRA = "DELETE*FROM CUENTA WHERE CODCLI='" + CODCLI + "' AND CODEMP='" + CODEMP + "' AND CODTCTA='" + CODTC + "' AND CODMON='" + MONEDA + "'"
       CN.Execute SQLBORRA
          RS.MoveLast
          RS.Requery
          FRADATOS.Enabled = False
          FRABOTONES.Enabled = True
          Call MOSTRAR
          Call HABILITARCONTROLES(True)
          CMDCONTROL(5).Enabled = False
       End If
       Case 3: CLAVE = InputBox("INGRESE CLAVE DEL SISTEMA", "SISTEMA BANCARIO")
       If UCase(CLAVE) = "SBLM" Then
          FRADATOS.Enabled = True
          FRABOTONES.Enabled = False
          Call HABILITARCONTROLES(False)
          CMDCONTROL(1).Enabled = False
          CMDCONTROL(5).Enabled = True
       End If
       Case 4: RS.CancelUpdate
               RS.MoveLast
               Call MOSTRAR
               FRADATOS.Enabled = False
               FRABOTONES.Enabled = True
               Call HABILITARCONTROLES(True)
               CMDCONTROL(5).Enabled = False
       Case 5:
          CODCLI = BUSCAR("CLIENTES", "NOMCLI", CBOCLI, 0)
          CODEMP = BUSCAR("EMPLEADOS", "EMPLEADO", CBOEMP, 0)
          CODTC = BUSCAR("TIPOCUENTA", "DESTCTA", CBOTC, 0)
          If OPTMON(0).Value = True Then MONEDA = "M-1"
          If OPTMON(1).Value = True Then MONEDA = "M-2"
          If OPTMON(2).Value = True Then MONEDA = "M-3"
          FRADATOS.Enabled = False
          FRABOTONES.Enabled = True
          CAMBIOS = "NROCTA='" + TXTNCTA + "',CODCLI='" + CODCLI + "',CODEMP='" + CODEMP + "',CODMON='" + MONEDA + "',CODTCTA='" + CODTC + "',MONTO= " + TXTMON + " ,FECHAREG=#" + TXTFREG + "#"
          SQLACT = "UPDATE CUENTA SET " + CAMBIOS + " WHERE NROCTA='" + TXTNCTA + "'"
          CN.Execute SQLACT
          RS.Requery
          Call HABILITARCONTROLES(True)
          CMDCONTROL(5).Enabled = False
          MsgBox ("REGISTRO ACTUALIZADO")
       Case 6:
       If MsgBox("¿DESEA SALIR DEL SISTEMA?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
          Unload Me
          MDISIS.Show
       End If
End Select
End Sub

Private Sub CMDMOVER_Click(Index As Integer)
Call MOVER(RS, Index)
Call MOSTRAR
End Sub

Private Sub Form_Load()
If RS.State = 1 Then RS.Close
Call COMBO("CLIENTES", CBOCLI, 1)
Call COMBO("TIPOCUENTA", CBOTC, 1)
Call COMBO("EMPLEADOS", CBOEMP, 1)
Dim SQL As String
SQL = "SELECT*FROM CUENTA"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Call MOSTRAR
Call HABILITARCONTROLES(True)
CMDCONTROL(5).Enabled = False
End Sub

Public Sub MOSTRAR()
Dim CODC As String
Dim CODTC As String
Dim CODM As String
Dim CODEMP As String
TXTNCTA = RS(0)
CODC = BUSCAR("CUENTA", "CODCLI", RS(1), 1)
CBOCLI = BUSCAR("CLIENTES", "CODCLI", CODC, 1)
CODEMP = BUSCAR("CUENTA", "CODEMP", RS(2), 2)
CBOEMP = BUSCAR("EMPLEADOS", "CODEMP", CODEMP, 1)
CODM = BUSCAR("CUENTA", "CODMON", RS(3), 3)
If CODM = "M-1" Then OPTMON(0).Value = True
If CODM = "M-2" Then OPTMON(1).Value = True
If CODM = "M-3" Then OPTMON(2).Value = True
CODTC = BUSCAR("CUENTA", "CODTCTA", RS(4), 4)
CBOTC = BUSCAR("TIPOCUENTA", "CODTCTA", CODTC, 1)
TXTMON = RS(5)
TXTFREG = RS(6)
End Sub

Public Sub HABILITARCONTROLES(HC As Boolean)
CMDCONTROL(0).Enabled = HC
CMDCONTROL(1).Enabled = Not HC
CMDCONTROL(2).Enabled = HC
CMDCONTROL(3).Enabled = HC
CMDCONTROL(4).Enabled = Not HC
CMDCONTROL(5).Enabled = HC
CMDCONTROL(6).Enabled = HC
End Sub

Public Sub GRABAR()
If OPTMON(0).Value = True Then MONEDA = "M-1"
If OPTMON(1).Value = True Then MONEDA = "M-2"
If OPTMON(2).Value = True Then MONEDA = "M-3"
CCLI = BUSCAR("CLIENTES", "NOMCLI", CBOCLI.Text, 0)
CTCTA = BUSCAR("TIPOCUENTA", "DESTCTA", CBOTC.Text, 0)
CEMP = BUSCAR("EMPLEADOS", "EMPLEADO", CBOEMP.Text, 0)
CAMPOS = "(NROCTA,CODCLI,CODEMP,CODMON,CODTCTA,MONTO,FECHAREG)"
VALORES = "('" + TXTNCTA + "','" + CCLI + "','" + CEMP + "','" + MONEDA + "','" + CTCTA + "', " + TXTMON + " ,#" + TXTFREG + "#)"
XGRABA = "INSERT INTO CUENTA " + CAMPOS + " VALUES " + VALORES
CN.Execute XGRABA
MsgBox ("NUEVA CUENTA CREADA")
End Sub

Public Sub GENERACODIGO()
Dim COD As String
Dim VAL As Integer
   COD = Mid(RS(0), 4, 2)
   VAL = COD + 1
   TXTNCTA = "M00" & Right(String(4 - Len(VAL), "0") + VAL, 4) & "-2005"
End Sub

