VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Captura"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2520
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4080
      TabIndex        =   6
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nº Vendedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   6210
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Command1_Click()



 VendedorNom(i) = Text1
 
 i = i + 1


End Sub


Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
End Sub

