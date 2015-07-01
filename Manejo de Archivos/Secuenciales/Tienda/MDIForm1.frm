VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Vera Smart Shop"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12765
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":2832
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7110
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Text            =   "Fecha:  "
            TextSave        =   "Fecha:  "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "04/03/2009"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Text            =   "Hora:  "
            TextSave        =   "Hora:  "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "1:53"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuArticulos 
         Caption         =   "Articulos"
      End
      Begin VB.Menu mnuVendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnuVentas 
         Caption         =   "Ventas"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuDiseñador 
      Caption         =   "Diseñador"
   End
   Begin VB.Menu mnuRep 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuCajeros 
         Caption         =   "&Cajeros"
      End
      Begin VB.Menu mnuArti 
         Caption         =   "&Articulos"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Click()
If RepVisible = True Then
frmAbout.Show

End If

End Sub

Private Sub mnuArti_Click()
frmRepArt.Show
End Sub

Private Sub mnuArticulos_Click()
frmArticulos.Show
End Sub

Private Sub mnuCajeros_Click()
frmRepCajeros.Show
End Sub

Private Sub mnuDiseñador_Click()
frmAbout.Show
End Sub

Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuVendedores_Click()
frmCajeros.Show
End Sub

Private Sub mnuVentas_Click()
frmLogin.Show
End Sub

Private Sub mnuVentasArticulos_Click()
Form4.Show
End Sub

Private Sub mnuVentasVendedores_Click()
Form5.Show
End Sub
