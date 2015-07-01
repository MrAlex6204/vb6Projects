VERSION 5.00
Begin VB.Form frmmenu 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VeraSoft Development"
   ClientHeight    =   7020
   ClientLeft      =   1020
   ClientTop       =   1215
   ClientWidth     =   9795
   Icon            =   "frmmenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmmenu.frx":954A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9840
   End
   Begin VB.Menu mnuaccesos 
      Caption         =   "&Accesos"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuarchivos 
      Caption         =   "A&rchivos"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "A&bout"
   End
   Begin VB.Menu mnusalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuaccesos_Click()
frmaccesos.Show
End Sub

Private Sub mnuarchivos_Click()
frmarchivos.Show
End Sub

Private Sub mnusalir_Click()
Unload Me
End Sub
