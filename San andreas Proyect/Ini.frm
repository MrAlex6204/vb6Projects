VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Ini 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11145
   FillColor       =   &H000040C0&
   Icon            =   "Ini.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10680
      Top             =   4320
   End
   Begin ComctlLib.ProgressBar progreso 
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   6600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   1,00000e5
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "VeraSoft Development"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Image Image1 
      DragIcon        =   "Ini.frx":954A
      Height          =   11520
      Left            =   -2160
      Picture         =   "Ini.frx":12A94
      Top             =   -1200
      Width           =   15360
   End
End
Attribute VB_Name = "Ini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo cargar
Exit Sub
cargar:
 FileCopy App.Path + "\COMDLG32.OCX", "C:\Windows\System32\COMDLG32.OCX"
 FileCopy App.Path + "\COMCTL32.OCX", "C:\Windows\System32\COMCTL32.OCX"
End Sub

Private Sub Timer1_Timer()
Dim i As Long
i = 0

While i < 100000
progreso.Value = i
i = i + 1
Wend
Unload Me
setup.Show
End Sub
