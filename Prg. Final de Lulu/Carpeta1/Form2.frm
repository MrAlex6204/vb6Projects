VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Captura de Vendedores"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   495
         Left            =   3360
         TabIndex        =   1
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   900
      End
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

Text1 = Empty
Text2 = Empty
End Sub



Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
End Sub
