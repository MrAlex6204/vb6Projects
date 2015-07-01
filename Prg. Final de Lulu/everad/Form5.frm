VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
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
      Height          =   3375
      Left            =   1320
      ScaleHeight     =   3315
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   1920
      Width           =   11535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Reporte x Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   4260
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Picture1_Click()
Dim i As Integer
i = 0
While i < 10
Picture1.Print VentasVendedores(i)

i = i + 1
Wend
End Sub
