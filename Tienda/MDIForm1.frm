VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Tienda"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9765
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":08D2
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu Cobranza 
      Caption         =   "Cobranza"
      Begin VB.Menu Cobrar 
         Caption         =   "Cobrar"
      End
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu Alta 
      Caption         =   "Alta"
      Begin VB.Menu Cajero 
         Caption         =   "Cajero"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRepLisArt 
         Caption         =   "Lista de Articulos"
      End
      Begin VB.Menu menuRepVenDia 
         Caption         =   "Venta Del Dia"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cajero_Click()
Form2.Show
End Sub

Private Sub Cerrar_Click()
End
End Sub

Private Sub Cobrar_Click()
frmLogin.Show
End Sub

Private Sub MDIForm_Load()
'Esta Parte es Para que Nuestra Coneccion Del Cual se Van a Tomar _
Los Datos Para el Reporte Se Conecte Con La Base de Datos Que esta En el Mismo _
Directorio de Nuestra Aplicacion

ConeccionDatos.Reportes.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
+ App.Path + "\Tienda.mdb;Persist Security Info=False"
End Sub

Private Sub menuRepVenDia_Click()
DataReport2.Show
End Sub

Private Sub mnuRepLisArt_Click()
DataReport1.Show
End Sub
