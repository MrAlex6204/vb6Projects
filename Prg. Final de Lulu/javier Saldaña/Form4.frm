VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Reporte de Articulos"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9930
   LinkTopic       =   "Form4"
   ScaleHeight     =   11085
   ScaleWidth      =   15240
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
      Height          =   5295
      Left            =   1320
      ScaleHeight     =   5235
      ScaleWidth      =   12075
      TabIndex        =   1
      Top             =   1800
      Width           =   12135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   555
      Left            =   6720
      TabIndex        =   0
      Top             =   840
      Width           =   1860
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
