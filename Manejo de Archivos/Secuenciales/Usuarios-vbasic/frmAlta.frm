VERSION 5.00
Begin VB.Form frmAlta 
   BackColor       =   &H00000000&
   Caption         =   "Alta"
   ClientHeight    =   5235
   ClientLeft      =   2460
   ClientTop       =   1950
   ClientWidth     =   7785
   Icon            =   "frmAlta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   7785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Alta de Usuarios"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2895
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   9255
      Begin VB.CommandButton Command3 
         Caption         =   "C&errar"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtApe 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtDirec 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtTel 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblApellidos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   1080
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   3960
      Picture         =   "frmAlta.frx":954A
      Top             =   240
      Width           =   4845
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Cerrar"
      Begin VB.Menu mnuCerrar 
         Caption         =   "&Cerrar"
      End
   End
End
Attribute VB_Name = "frmAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
ArchUsuarios = FreeFile

If txtNombre = Empty Or txtApe = Empty Or txtDirec = Empty Or txtTel = Empty Then
MsgBox "Asegurese de Llenar Todos los Capos", vbCritical, "VeraSoft Developmet"
Exit Sub
End If

Open App.path + "\ArchMaster.Dat" For Append As #ArchUsuarios
Print #ArchUsuarios, UCase(txtNombre)
Print #ArchUsuarios, UCase(txtApe)
Print #ArchUsuarios, UCase(txtDirec)
Print #ArchUsuarios, UCase(txtTel)
Close #ArchUsuarios
MsgBox "Registro Agregado", vbExclamation, "VeraSoft Development"

txtNombre = Empty
txtApe = Empty
txtDirec = Empty
txtTel = Empty


End Sub

Private Sub Command2_Click()
txtNombre = Empty
txtApe = Empty
txtDirec = Empty
txtTel = Empty
End Sub





Private Sub Command3_Click()
Unload Me
End Sub

Private Sub mnuCerrar_Click()
Unload Me
End Sub
