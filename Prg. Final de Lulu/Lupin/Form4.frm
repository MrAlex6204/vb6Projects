VERSION 5.00
Begin VB.Form Form4 
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
      Height          =   3375
      Left            =   1320
      ScaleHeight     =   3315
      ScaleWidth      =   10275
      TabIndex        =   1
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Reportes de Articulos"
      BeginProperty Font 
         Name            =   "CityBlueprint"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   840
      Width           =   3690
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
