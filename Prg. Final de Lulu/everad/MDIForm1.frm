VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "Everad"
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
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuVentasArticulos 
         Caption         =   "Ventas Articulos"
      End
      Begin VB.Menu mnuVentasVendedores 
         Caption         =   "Ventas Vendedores"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub mnuVentasArticulos_Click()
Form4.Show
End Sub

Private Sub mnuVentasVendedores_Click()
Form5.Show
End Sub
