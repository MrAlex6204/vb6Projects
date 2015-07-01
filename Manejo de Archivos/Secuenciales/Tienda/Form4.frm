VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte x Articulo"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9930
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9930
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4380
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   11535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte de Articulos Vendidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   6480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Paint()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim cadena As String
Dim i, num As Integer
i = 0

While i < 10


num = VentasArticulos(i)
If num = 0 Then

Else
cadena = Str(num) + "--->" + NomArticulo(num - 1)
List1.AddItem cadena
End If


i = i + 1
Wend

End Sub
