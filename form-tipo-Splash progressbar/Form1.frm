VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar form"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   480
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmSplash.Show
Form1.Hide
End Sub

Private Sub Command2_Click()
End
End Sub
