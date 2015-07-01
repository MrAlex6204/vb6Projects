VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Reporte x Articulo"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9930
   LinkTopic       =   "Form4"
   ScaleHeight     =   4320
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   15075
      TabIndex        =   0
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Paint()
Dim i As Integer
i = 0
While i < 10
Picture1.Print VentasArticulos(i)

i = i + 1
Wend
End Sub
