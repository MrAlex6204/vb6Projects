VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reporte x Vendedor"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   4680
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
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   12735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reporte de Ventas por Vendedor"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   7035
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Dim cadena As String
Dim i, num As Integer
i = 0

While i < 10


num = VentasVendedores(i)
If num = 0 Then

Else
cadena = Str(num) + "--->" + VendedorNom(num - 1)
List1.AddItem cadena
End If


i = i + 1
Wend
End Sub

