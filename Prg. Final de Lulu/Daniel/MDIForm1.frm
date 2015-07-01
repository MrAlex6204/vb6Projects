VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "MDIForm1"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   9675
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Begin VB.Menu mnuVentasArticulos 
         Caption         =   "Reporte Ventas Articulos"
      End
      Begin VB.Menu mnuVentasVendedores 
         Caption         =   "Reporte Ventas Vendedores"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

End Sub

Private Sub mnuArticulos_Click()
Form1.Show
End Sub

Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuVendedores_Click()
Form2.Show
End Sub

Private Sub mnuVentas_Click()
Form3.Show
End Sub

