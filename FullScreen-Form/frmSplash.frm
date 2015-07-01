VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   10830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   2040
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5130
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   8520
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   4680
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VeraSoft Develoment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   3000
         TabIndex        =   3
         Top             =   4200
         Width           =   3600
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VeraSoft Develoment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   2040
         Width           =   3600
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me

End Sub

Private Sub Form_Load()
'------------------------------
'para poner el form en pantalla completa
' al atamaño de la resolucion  del monitor
frmSplash.Top = 0
frmSplash.Left = 0
frmSplash.Width = Screen.Width
frmSplash.Height = Screen.Height
Frame1.Top = Screen.Height / 4
Frame1.Left = Screen.Width / 4
'-------------------------------
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()
End Sub

Sub ProgressBar1_Click()
ProgressBar1.Visible = True
Dim i, x As Integer
i = 1
ProgressBar1.Max = 100000


While i < 100000

Label1.Caption = Str((i * 100) / ProgressBar1.Max)
i = i + 1
ProgressBar1.Value = i

Wend
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
Call ProgressBar1_Click


Timer2.Enabled = False
End Sub

