VERSION 5.00
Begin VB.Form FRMCLI 
   BackColor       =   &H00004000&
   Caption         =   "DATOS GENERALES DE CLIENTES"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   7155
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDCONTROL 
      Height          =   615
      Index           =   6
      Left            =   6000
      Picture         =   "FORM1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "SALIR"
      Top             =   5150
      Width           =   975
   End
   Begin VB.Frame FRADATOS 
      BackColor       =   &H00004000&
      Enabled         =   0   'False
      Height          =   3255
      Left            =   1320
      TabIndex        =   12
      Top             =   600
      Width           =   4575
      Begin VB.ComboBox CBODIS 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00004000&
         Caption         =   "SEXO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000CF3BA&
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   3015
         Begin VB.OptionButton OPTSEXO 
            BackColor       =   &H00004000&
            Caption         =   "FEMENINO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000CF3BA&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OPTSEXO 
            BackColor       =   &H00004000&
            Caption         =   "MASCULINO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000CF3BA&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox TXTTEL 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TXTDIR 
         Height          =   615
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TXTNOM 
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TXTCOD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRITO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000AFAB8&
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         Caption         =   "TELEFONO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000CF3BA&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         Caption         =   "DIRECCION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000CF3BA&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         Caption         =   "NOMBRE DEL CLIENTE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000CF3BA&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         Caption         =   "CODIGO DE CLIENTE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000CF3BA&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   6975
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   5
         Left            =   3960
         Picture         =   "FORM1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "ACTUALIZAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   4
         Left            =   4920
         Picture         =   "FORM1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "CANCELAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   3
         Left            =   3000
         Picture         =   "FORM1.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "MODIFICAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   2
         Left            =   2040
         Picture         =   "FORM1.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "BORRAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   1
         Left            =   1080
         Picture         =   "FORM1.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "GRABAR"
         Top             =   230
         Width           =   975
      End
      Begin VB.CommandButton CMDCONTROL 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "FORM1.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "NUEVO"
         Top             =   230
         Width           =   975
      End
   End
   Begin VB.Frame FRABOTONES 
      BackColor       =   &H00004000&
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   3960
      Width           =   5055
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   3
         Left            =   3720
         Picture         =   "FORM1.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   2
         Left            =   2520
         Picture         =   "FORM1.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   1
         Left            =   1320
         Picture         =   "FORM1.frx":2A94
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton CMDMOVER 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "FORM1.frx":2ED6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   120
      Picture         =   "FORM1.frx":3318
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE CLIENTES ASOCIADOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CF3BA&
      Height          =   345
      Left            =   1080
      TabIndex        =   11
      Top             =   120
      Width           =   5100
   End
End
Attribute VB_Name = "FRMCLI"
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
               TXTCOD.Enabled = False
               Call HABILITARCONTROLES(False)
       Case 1: Call GRABAR
               FRADATOS.Enabled = False
               RS.Requery
               Call HABILITARCONTROLES(True)
               FRABOTONES.Enabled = True
               CMDCONTROL(5).Enabled = False
       Case 2:
       Dim CODDIS As String
       If MsgBox("¿DESEA BORRAR EL SGTE. REGISTRO?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
       CODDIS = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
          SQLBORRA = "DELETE*FROM CLIENTES WHERE CODDIS='" + CODDIS + "'"
          CN.Execute SQLBORRA
          RS.MoveLast
          RS.Requery
          FRADATOS.Enabled = False
          FRABOTONES.Enabled = True
          Call MOSTRAR
          Call HABILITARCONTROLES(True)
          CMDCONTROL(5).Enabled = False
       End If
       Case 3: CLAVE = InputBox("INGRESAR CLAVE DEL SISTEMA", "SISTEMA BANACARIO")
       If UCase(CLAVE) = "SBLM" Then
          FRADATOS.Enabled = True
          FRABOTONES.Enabled = False
          Call HABILITARCONTROLES(False)
          CMDCONTROL(5).Enabled = True
          CMDCONTROL(1).Enabled = False
       End If
       Case 4: RS.CancelUpdate
               FRADATOS.Enabled = False
               FRABOTONES.Enabled = True
               RS.MoveLast
               Call MOSTRAR
               Call HABILITARCONTROLES(True)
               CMDCONTROL(5).Enabled = False
       Case 5:
          Dim SEX As String
          CODD = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
          If OPTSEXO(0).Value = True Then SEX = "M"
          If OPTSEXO(1).Value = True Then SEX = "F"
          FRADATOS.Enabled = False
          FRABOTONES.Enabled = True
          CAMBIOS = "NOMCLI='" + TXTNOM + "',DIRCLI='" + TXTDIR + "',TELCLI='" + TXTTEL + "',SEXO='" + SEX + "',CODDIS='" + CODD + "'"
          SQLACT = "UPDATE CLIENTES SET " + CAMBIOS + " WHERE CODCLI='" + TXTCOD + "'"
          CN.Execute SQLACT
          RS.Requery
          Call HABILITARCONTROLES(True)
          CMDCONTROL(5).Enabled = False
          MsgBox ("REGISTRO ACTUALIZADO")
       Case 6:
       If MsgBox("¿DESEA SALIR?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
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
Call COMBO("DISTRITOS", CBODIS, 1)
Dim SQL As String
SQL = "SELECT*FROM CLIENTES"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Call MOSTRAR
Call HABILITARCONTROLES(True)
CMDCONTROL(5).Enabled = False
End Sub

Public Sub MOSTRAR()
Dim CODDIS As String
TXTCOD = RS(0)
TXTNOM = RS(1)
TXTDIR = RS(2)
TXTTEL = RS(3)
SEXO = RS(4)
If SEXO = "M" Then OPTSEXO(0).Value = True
If SEXO = "F" Then OPTSEXO(1).Value = True
CODDIS = BUSCAR("CLIENTES", "CODDIS", RS(5), 5)
CBODIS = BUSCAR("DISTRITOS", "CODDIS", CODDIS, 1)
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
Dim CODDIS As String
If OPTSEXO(0).Value = True Then SEXO = "M"
If OPTSEXO(1).Value = True Then SEXO = "F"
CODDIS = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
CAMPOS = "(CODCLI,NOMCLI,DIRCLI,TELCLI,SEXO,CODDIS)"
VALORES = "('" + TXTCOD + "','" + TXTNOM + "','" + TXTDIR + "','" + TXTTEL + "','" + SEXO + "','" + CODDIS + "')"
XGRABA = "INSERT INTO CLIENTES " + CAMPOS + " VALUES " + VALORES
CN.Execute XGRABA
MsgBox ("DATOS DE CLIENTE ACEPTADO")
End Sub

Public Sub GENERACODIGO()
Dim COD As String
Dim VAL As Integer
   COD = Mid(RS(0), 4, 2)
   VAL = COD + 1
   TXTCOD = "C00" & Right(String(4 - Len(VAL), "0") + VAL, 4) & "-2005"
End Sub
