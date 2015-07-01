VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "VeraSoft"
   ClientHeight    =   3795
   ClientLeft      =   2295
   ClientTop       =   1575
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2619.376
   ScaleMode       =   0  'User
   ScaleWidth      =   9057.153
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   2655
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAbout.frx":0000
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   7920
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   0
      Picture         =   "frmAbout.frx":007A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
