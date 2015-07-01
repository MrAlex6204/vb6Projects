VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte x Vendedor"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   11085
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
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
      Height          =   4335
      Left            =   360
      ScaleHeight     =   4275
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   4365
      Left            =   8400
      Picture         =   "Form5.frx":0000
      Top             =   1800
      Width           =   4140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   1920
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
Dim i As Integer
i = 0
While i < 10
Picture1.Print VentasVendedores(i)

i = i + 1
Wend
End Sub
