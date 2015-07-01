VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Kill All"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3300
      Top             =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Forms in Array:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   1020
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmWindows() As New Form2

Private Sub Command1_Click()
    UpdateWindowList frmWindows
    frmWindows(UBound(frmWindows)).Text1 = UBound(frmWindows)
    frmWindows(UBound(frmWindows)).Show
    ReDim Preserve frmWindows(UBound(frmWindows) + 1)
End Sub

Private Sub Command2_Click()
    KillWindows frmWindows
End Sub

Private Sub Form_Load()
    ReDim frmWindows(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    KillWindows frmWindows
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = UBound(frmWindows)
End Sub

