VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   1920
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const KEYEVENTF_KEYUP = &H2
Const KEYEVENTF_EXTENDEDKEY = &H1
  
  
'Declaración del Api keybd_event para la presión de tecla
  
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)
  
  
  
Sub Pulsar_Tecla(Tecla As Long)
  
    Call keybd_event(Tecla, 0, 0, 0)
  
    Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
  
End Sub
  
Private Sub Command4_Click()
Timer1.Enabled = False
End Sub

' Presiona la tecla Tab
Private Sub Timer1_Timer()
Call Pulsar_Tecla(vbKeyTab)
End Sub


