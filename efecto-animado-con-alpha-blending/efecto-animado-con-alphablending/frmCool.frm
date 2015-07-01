VERSION 5.00
Begin VB.Form frmCool 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " http://www.i.com.ua/~aka"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmCool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCool.frx":030A
   ScaleHeight     =   8310
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   120
      Picture         =   "frmCool.frx":2C26C
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   4530
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   120
      Picture         =   "frmCool.frx":581CE
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   4530
   End
End
Attribute VB_Name = "frmCool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
 Run_Blending Picture1, picSrc
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End ' STOP Do Loop
End Sub

