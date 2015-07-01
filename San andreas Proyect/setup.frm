VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form setup 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14250
   LinkTopic       =   "Form2"
   Picture         =   "setup.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   7440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   10680
      TabIndex        =   2
      Top             =   8880
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   12120
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Carpeta de Intalacion"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Instalar"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Intalar Crack"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Espere Mientras se Intala San andreas Para Despues Instalar el Crack....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   7200
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim Path As String
    
Path = DLG_BrowseFolder(Me.hwnd, "Seleccione directorio de Intalacion")
    
If Path = Empty Then
MsgBox "Seleccione el Path de San Andreas", vbCritical, "Vera Soft Development"
Else
Label1.Visible = True
FileCopy App.Path + "\crack_para_poder_guardar_las_partidas_gta_sa.exe", Path + "\gta_sa.exe"
End If

End Sub

Private Sub Command2_Click()
Shell App.Path + "\GTA San Andreas RIP.exe", vbNormalFocus
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Me.WindowState = 2
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub Timer1_Timer()
Dim i As Long
i = 0
While i < 1000000
i = i + 1
Wend
Command1.Enabled = True
Label1.Caption = "Intale el Crack"
Timer1.Enabled = False
End Sub
