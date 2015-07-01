VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   8010
   ClientTop       =   9540
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   6
      Left            =   360
      Top             =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   945
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
frmSplash1.Height = 0
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
i = frmSplash1.Height


i = frmSplash1.Height
While i > 0
frmSplash1.Height = i
i = i - 1

Wend


Unload Me

End Sub

Sub Label1_Click()
Timer1.Enabled = False

Dim i As Integer
i = 0
While i < 10000
i = i + 1
Wend




End Sub

Private Sub Timer1_Timer()
Dim i As Integer
While i < 2085
frmSplash1.Height = i
i = i + 1

Wend
Label1.Caption = "Conectado:" + Form1.User
Label1_Click
End Sub
