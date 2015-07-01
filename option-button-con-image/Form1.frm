VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1485
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.Image ImgHiden 
      Height          =   480
      Index           =   1
      Left            =   0
      Top             =   480
      Width           =   3000
   End
   Begin VB.Image ImgHiden 
      Height          =   480
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":0A36
      Top             =   520
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":11FA
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opcion As Integer

Private Sub Form_Load()
Image1(1).Visible = True
Image1(0).Visible = False
End Sub

Private Sub ImgHiden_Click(Index As Integer)
Image1(opcion).Visible = False
Image1(Index).Visible = True


opcion = Index
End Sub
