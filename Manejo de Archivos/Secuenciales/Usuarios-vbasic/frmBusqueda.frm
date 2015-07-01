VERSION 5.00
Begin VB.Form frmBusqueda 
   BackColor       =   &H00000000&
   Caption         =   "Busqueda"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "frmBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Busqueda"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "C&errar"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtBuscar 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame FrmeDatos 
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
      Left            =   4800
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   6855
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
         TabIndex        =   15
         Top             =   360
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
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
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
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
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
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
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
         TabIndex        =   6
         Top             =   1320
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
         TabIndex        =   5
         Top             =   840
         Width           =   1350
      End
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
         TabIndex        =   4
         Top             =   360
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
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   11895
   End
   Begin VB.Label lblCantidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empty"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados de la Busqueda:"
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
      TabIndex        =   1
      Top             =   4560
      Width           =   3510
   End
   Begin VB.Image Image1 
      Height          =   2490
      Left            =   3000
      Picture         =   "frmBusqueda.frx":954A
      Top             =   0
      Width           =   5070
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()
ListUsuarios.Clear
Dim Contenido As String
Dim Encontrado As Boolean
Dim Cantidad As Integer
Cantidad = 0
Encontrado = False
ArchUsuarios = FreeFile
If Verificar_Existe(App.path + "\ArchMaster.Dat") = True Then

Open App.path + "\ArchMaster.Dat" For Input As #ArchUsuarios
Do While Not EOF(ArchUsuarios)
        
        'Lee la linea
        Line Input #ArchUsuarios, Contenido
        
        
        If UCase(txtBuscar) = Contenido Then
            Encontrado = True
            Cantidad = Cantidad + 1
            ListUsuarios.AddItem "Nombre(s):" + Contenido
            Line Input #ArchUsuarios, Contenido
            ListUsuarios.AddItem "   Apellidos:" + Contenido
            Line Input #ArchUsuarios, Contenido
            ListUsuarios.AddItem "   Direccion:" + Contenido
            Line Input #ArchUsuarios, Contenido
            ListUsuarios.AddItem "   Telefono :" + Contenido
            ListUsuarios.AddItem ""
        Else
            Line Input #ArchUsuarios, Contenido
            Line Input #ArchUsuarios, Contenido
            Line Input #ArchUsuarios, Contenido
            End If
        
Loop
Close #ArchUsuarios

If Encontrado = False Then
lblCantidad = "No se Encontraron Usuarios"
Else
lblCantidad = Cantidad
End If

Else
MsgBox "No se Encontro el Archivo: ArchMaster.Dat", vbCritical, "VeraSoft Development"
Unload Me
Me.Hide
End If

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblCantidad = ""
ListUsuarios.Appearance = 0

End Sub

Private Sub ListUsuarios_Click()

Dim cadena As String
Dim i As Integer
If ListUsuarios.ListIndex <> -1 Then
   'para poner el dato de un ittem seleccionado
  cadena = Mid(ListUsuarios.List(ListUsuarios.ListIndex), 1, 10)
  
If cadena = "Nombre(s):" Then

  FrmeDatos.Visible = True
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

Private Sub menuCerrarBus_Click()
Unload Me
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
