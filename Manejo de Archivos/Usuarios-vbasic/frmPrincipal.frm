VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmPrincipal 
   BackColor       =   &H00000000&
   Caption         =   "Manejo de Archivos"
   ClientHeight    =   4860
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7230
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmPrincipal.frx":954A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Text            =   "Hora:"
            TextSave        =   "Hora:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "10:51 p.m."
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Text            =   "Fecha:"
            TextSave        =   "Fecha:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2858
            TextSave        =   "24/02/2009"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNvo 
         Caption         =   "&Nuevo"
         Begin VB.Menu mnuUsuario 
            Caption         =   "&Usuario"
         End
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuraciones"
         Begin VB.Menu mnuELiminararch 
            Caption         =   "&Eliminar Archivos"
         End
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuELiminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar Todos"
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAlta_Click()
frmAlta.Show
End Sub

Private Sub mnuBuscar_Click()
frmBusqueda.Show
End Sub

Private Sub mnuELiminar_Click()
frmEliminar.Show

End Sub

Private Sub mnuELiminararch_Click()
If Verificar_Existe(App.path + "\ArchMaster.Dat") = True Then
Kill App.path + "\ArchMaster.Dat"
MsgBox "Archivo: ArchMaster.Dat Fue Eliminado", vbCritical, "VeraSoft Development"
Else
MsgBox "No se Encontro el Archivo: ArchMaster.Dat", vbCritical, "VeraSoft Development"
End If
End Sub

Private Sub mnuMostrar_Click()
frmMostrar.Show
End Sub

Private Sub mnuSalir_Click()
Unload frmAlta
Unload frmMostrar
End
End Sub

Private Sub mnuUsuario_Click()
frmAlta.Show
End Sub
