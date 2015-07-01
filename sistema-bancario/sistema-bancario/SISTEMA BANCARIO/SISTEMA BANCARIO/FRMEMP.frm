VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMEMP 
   BackColor       =   &H0024CCEA&
   Caption         =   "DATOS GENERALES DE EMPLEADOS"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTEMP 
      Height          =   6135
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   2411754
      TabCaption(0)   =   "REGISTRO DE EMPLEADOS"
      TabPicture(0)   =   "FRMEMP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FRACOD"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FRADATOS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FRABOTONES"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FRACONTROLES"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "BUSQUEDA POR CARGOS"
      TabPicture(1)   =   "FRMEMP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CBOBCAR"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MHFBCAR"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "REFERENCIAS DE CODIGOS"
      TabPicture(2)   =   "FRMEMP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MHFCAR"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFBCAR 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   35
         Top             =   1200
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7011
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox CBOBCAR 
         Height          =   315
         Left            =   -74040
         TabIndex        =   34
         Top             =   600
         Width           =   2055
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHFCAR 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5530
         _Version        =   393216
         FixedCols       =   0
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
      Begin VB.Frame FRACONTROLES 
         Height          =   4575
         Left            =   6240
         TabIndex        =   23
         Top             =   780
         Width           =   1215
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   5
            Left            =   120
            Picture         =   "FRMEMP.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "ACTUALIZAR"
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   6
            Left            =   120
            Picture         =   "FRMEMP.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "SALIR"
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   4
            Left            =   120
            Picture         =   "FRMEMP.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "CANCELAR"
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   3
            Left            =   120
            Picture         =   "FRMEMP.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "MODIFICAR"
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   2
            Left            =   120
            Picture         =   "FRMEMP.frx":115C
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "BORRAR"
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   1
            Left            =   120
            Picture         =   "FRMEMP.frx":159E
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "GRABAR"
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton CMDCONTROL 
            Height          =   615
            Index           =   0
            Left            =   120
            Picture         =   "FRMEMP.frx":19E0
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "NUEVO"
            Top             =   230
            Width           =   975
         End
      End
      Begin VB.Frame FRABOTONES 
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   4860
         Width           =   5055
         Begin VB.CommandButton CMDMOVER 
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "FRMEMP.frx":1E22
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   200
            Width           =   1215
         End
         Begin VB.CommandButton CMDMOVER 
            Height          =   495
            Index           =   1
            Left            =   1320
            Picture         =   "FRMEMP.frx":2264
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   200
            Width           =   1215
         End
         Begin VB.CommandButton CMDMOVER 
            Height          =   495
            Index           =   2
            Left            =   2520
            Picture         =   "FRMEMP.frx":26A6
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   200
            Width           =   1215
         End
         Begin VB.CommandButton CMDMOVER 
            Height          =   495
            Index           =   3
            Left            =   3720
            Picture         =   "FRMEMP.frx":2AE8
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   200
            Width           =   1215
         End
      End
      Begin VB.Frame FRADATOS 
         Enabled         =   0   'False
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   1620
         Width           =   4935
         Begin VB.ComboBox CBODIS 
            Height          =   315
            Left            =   1080
            TabIndex        =   17
            Top             =   2640
            Width           =   3135
         End
         Begin VB.Frame Frame2 
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
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   3015
            Begin VB.OptionButton OPTSEX 
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton OPTSEX 
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
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   14
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox TXTNOM 
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TXTTEL 
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox TXTDIR 
            Height          =   855
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   2640
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NOMBRE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   750
         End
      End
      Begin VB.Frame FRACOD 
         Caption         =   "CODIGOS:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   3975
         Begin VB.ComboBox CBOCAR 
            Height          =   315
            Left            =   2160
            TabIndex        =   31
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TXTCODEMP 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CODIGO DE CARGO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "CODIGO DE EMPLEADO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1980
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CARGO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   33
         Top             =   600
         Width           =   660
      End
   End
   Begin VB.Image Image1 
      Height          =   5580
      Left            =   9000
      Picture         =   "FRMEMP.frx":2F2A
      Top             =   1200
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE EMPLEADOS REGISTRADOS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FRMEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBOBCAR_Click()
Select Case CBOBCAR.ListIndex
       Case 0 To 3:
       Dim CADBUS As String
       If RS.State = 1 Then RS.Close
       CADBUS = BUSCAR("CARGOS", "DESCAR", CBOBCAR, 0)
       SQLBUS = "SELECT CODEMP,EMPLEADO,DIR,TEL,SEXO,D.DISTRITO FROM EMPLEADOS E, DISTRITOS D WHERE E.CODDIS=D.CODDIS AND CODCAR='" + CADBUS + "'"
       RS.Open SQLBUS, CN, adOpenDynamic, adLockOptimistic
       Set MHFBCAR.DataSource = RS
End Select
End Sub

Private Sub CMDCONTROL_Click(Index As Integer)
Select Case Index
       Case 0: RS.MoveLast
               Call LIMPIAR(Me)
               Call GENERACODIGO
               FRADATOS.Enabled = True
               FRACOD.Enabled = True
               FRABOTONES.Enabled = False
               Call HABILITARCONTROLES(False)
       Case 1: Call GRABAR
               FRADATOS.Enabled = False
               FRABOTONES.Enabled = True
               RS.Requery
               Call HABILITARCONTROLES(True)
               CMDCONTROL(5).Enabled = False
               FRACOD.Enabled = False
       Case 2:
       Dim CODDIS As String
       If MsgBox("¿DESEA BORRAR EL SGTE. REGISTRO?", vbYesNo, "SISTEMA BANCARIO") = vbYes Then
       CODDIS = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
       SQLBORRA = "DELETE*FROM EMPLEADOS WHERE CODEMP='" + TXTCODEMP + "'"
       CN.Execute SQLBORRA
       RS.MoveLast
       RS.Requery
       FRADATOS.Enabled = False
       FRABOTONES.Enabled = True
       FRACOD.Enabled = False
       Call MOSTRAR
       Call HABILITARCONTROLES(True)
       CMDCONTROL(5).Enabled = False
       End If
       Case 3: CLAVE = InputBox("INGRESE CLAVE DEL SISTEMA", "SISTEMA BANCARIO")
       If UCase(CLAVE) = "SBLM" Then
          FRADATOS.Enabled = True
          FRABOTONES.Enabled = False
          FRACOD.Enabled = True
          Call HABILITARCONTROLES(False)
          CMDCONTROL(5).Enabled = True
          CMDCONTROL(1).Enabled = False
       End If
       Case 4: RS.CancelUpdate
               RS.MoveLast
               Call MOSTRAR
               FRADATOS.Enabled = False
               FRABOTONES.Enabled = True
               FRACOD.Enabled = False
               Call HABILITARCONTROLES(True)
               CMDCONTROL(5).Enabled = False
       Case 5:
       CDIST = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
       If OPTSEX(0).Value = True Then SEXO = "M"
       If OPTSEX(1).Value = True Then SEXO = "F"
       FRADATOS.Enabled = False
       FRABOTONES.Enabled = True
       FRACOD.Enabled = False
       CAMBIOS = "CODCAR='" + CBOCAR + "',EMPLEADO='" + TXTNOM + "',DIR='" + TXTDIR + "',TEL='" + TXTTEL + "',SEXO='" + SEXO + "',CODDIS='" + CDIST + "'"
       SQLACT = "UPDATE EMPLEADOS SET " + CAMBIOS + " WHERE CODEMP='" + TXTCODEMP + "'"
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
Call COMBO("CARGOS", CBOCAR, 0)
 Call COMBO("CARGOS", CBOBCAR, 1)
SQL = "SELECT*FROM EMPLEADOS"
RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Call MOSTRAR
Call HABILITARCONTROLES(True)
CMDCONTROL(5).Enabled = False
SSTEMP.Tab = 0
End Sub

Private Sub SSTEMP_Click(PreviousTab As Integer)
Select Case SSTEMP.Tab
       Case 0:
       If RS.State = 1 Then RS.Close
       SQL = "SELECT*FROM EMPLEADOS"
       RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
       Call MOSTRAR
       Case 2:
       If RS.State = 1 Then RS.Close
       SQLC = "SELECT CODCAR AS [CODIGO],DESCAR AS [CARGO] FROM CARGOS"
       RS.Open SQLC, CN, adOpenDynamic, adLockOptimistic
       Set MHFCAR.DataSource = RS
       MHFCAR.ColWidth(0) = 1000
       MHFCAR.ColWidth(1) = 2300
End Select
End Sub

Public Sub MOSTRAR()
Dim CDIST As String
TXTCODEMP = RS(0)
TXTNOM = RS(1)
CBOCAR = RS(2)
TXTDIR = RS(3)
TXTTEL = RS(4)
SEXO = RS(5)
If SEXO = "M" Then OPTSEX(0).Value = True
If SEXO = "F" Then OPTSEX(1).Value = True
CDIST = BUSCAR("EMPLEADOS", "CODDIS", RS(6), 6)
CBODIS = BUSCAR("DISTRITOS", "CODDIS", CDIST, 1)
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

Public Sub GENERACODIGO()
Dim COD As String
Dim VAL As Integer
   COD = Mid(RS(0), 6, 2)
   VAL = COD + 1
   TXTCODEMP = "EMP000" & Right(String(4 - Len(VAL), "0") + VAL, 4)
End Sub

Public Sub GRABAR()
Dim CDIST As String
If OPTSEX(0).Value = True Then SEXO = "M"
If OPTSEX(1).Value = True Then SEXO = "F"
CDIST = BUSCAR("DISTRITOS", "DISTRITO", CBODIS.Text, 0)
CAMPOS = "(CODEMP,EMPLEADO,CODCAR,DIR,TEL,SEXO,CODDIS)"
VALORES = "('" + TXTCODEMP + "','" + TXTNOM + "','" + CBOCAR + "','" + TXTDIR + "','" + TXTTEL + "','" + SEXO + "','" + CDIST + "')"
XGRABA = "INSERT INTO EMPLEADOS " + CAMPOS + " VALUES " + VALORES
CN.Execute XGRABA
MsgBox ("DATOS DE EMPLEADO REGISTRADOS")
End Sub
