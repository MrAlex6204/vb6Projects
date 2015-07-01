VERSION 5.00
Begin VB.MDIForm MDIProyecto 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Menu de Aplicaciones"
   ClientHeight    =   10710
   ClientLeft      =   570
   ClientTop       =   360
   ClientWidth     =   14385
   Icon            =   "MDIProyecto.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIProyecto.frx":030A
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuColores 
         Caption         =   "&Colores"
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu mnuTemperatura 
         Caption         =   "&Temperatura"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
   End
End
Attribute VB_Name = "MDIProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCalculadora_Click()
frmMinical.Show
End Sub

Private Sub mnuColores_Click()
frmColores.Show
End Sub

Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuTemperatura_Click()
frmTemp.Show
End Sub
