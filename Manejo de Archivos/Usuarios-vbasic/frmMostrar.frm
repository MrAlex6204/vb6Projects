VERSION 5.00
Begin VB.Form frmMostrar 
   BackColor       =   &H00000000&
   Caption         =   "Usuarios Del Sistema"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmMostrar.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame frmeMostrar 
      BackColor       =   &H00000000&
      Caption         =   "Datos"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono.:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label lbltel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbldirec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblape 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.ListBox ListUsuarios 
      BackColor       =   &H00000000&
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
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   11655
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   3960
      Picture         =   "frmMostrar.frx":954A
      Top             =   -120
      Width           =   4845
   End
   Begin VB.Menu mnuArch 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuCerrarmost 
         Caption         =   "&Cancelar"
      End
   End
End
Attribute VB_Name = "frmMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ListUsuarios.Appearance = 0
ArchUsuarios = FreeFile


If Verificar_Existe(App.path + "\ArchMaster.Dat") = True Then

Open App.path + "\ArchMaster.Dat" For Input As #ArchUsuarios

Dim Contenido As String
Do While Not EOF(ArchUsuarios)
        
        'Lee la linea
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "Nombre(s):" + Contenido
        
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Apellidos:" + Contenido
        
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Direccion:" + Contenido
        
         Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Telefono :" + Contenido
        
        ListUsuarios.AddItem ""
Loop
Close #ArchUsuarios

Else
MsgBox "No se Encontro el Archivo: ArchMaster.Dat", vbCritical, "VeraSoft Development"
Me.Hide
End If




End Sub

Private Sub ListUsuarios_Click()
frmeMostrar.Visible = True
Dim cadena As String
Dim i As Integer
If ListUsuarios.ListIndex <> -1 Then
   'para poner el dato de un ittem seleccionado
  cadena = Mid(ListUsuarios.List(ListUsuarios.ListIndex), 1, 10)
  If cadena = "Nombre(s):" Then
  lblnom.Visible = True
  lblape.Visible = True
  lbldirec.Visible = True
  lbltel.Visible = True
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex))
  lblnom = Mid(ListUsuarios.List(ListUsuarios.ListIndex), 11, (i - 10 + 1))
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 1))
  lblape = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 1), 14, (i - 10 + 1))
  
   i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 2))
  lbldirec = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 2), 14, (i - 10 + 1))
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 3))
  lbltel = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 3), 14, (i - 10 + 1))
  
  End If

End If
End Sub

Private Sub mnuCerrarmost_Click()
Unload Me
End Sub
