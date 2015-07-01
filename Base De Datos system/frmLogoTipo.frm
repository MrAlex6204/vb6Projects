VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5400
   ClientLeft      =   3480
   ClientTop       =   3480
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6600
      Top             =   240
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   0
      Max             =   10000
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Universidad Autonoma de Tamaulipas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   6555
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "VeraSoftWare Develoment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright:Oscar A. Vera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6360
      TabIndex        =   1
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   0
      Picture         =   "frmLogoTipo.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
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

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()

End Sub

Private Sub Form_Load()
Transparent.Aplicar_Transparencia Me.hWnd, 215
End Sub

Sub ProgressBar1_Click()
Timer1.Enabled = False
Dim i As Integer
i = 0
While i < ProgressBar1.Max
ProgressBar1.Value = i
i = i + 1
Wend
 Unload Me
MDIForm1.Show

End Sub

Private Sub Timer1_Timer()
Call ProgressBar1_Click
End Sub
