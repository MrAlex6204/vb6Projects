VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "&Abrir"
   ClientHeight    =   9120
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   13995
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FileManager.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox unidad 
      BackColor       =   &H000040C0&
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DirListBox Directorio 
      BackColor       =   &H000040C0&
      Height          =   5040
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File 
      BackColor       =   &H000040C0&
      Height          =   5160
      Left            =   5280
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Menu OpenFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu Open 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu Salir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo siguiente ' si produce un error se pasa hasta donde

'dice siguiente

If Text1.Text = Empty Then
MsgBox "Telclee el Path de un Exe", vbCritical, "Teclee La direccion Correcta"

siguiente: 'Apartir de aqui se ejecuta si huvo un error en el prg
Text1.SetFocus
Else
Shell Text1.Text

End If


End Sub

Private Sub File1_Click()

End Sub

Private Sub Drive1_Change()

End Sub

Private Sub Directorio_Change()
File.Path = Directorio.Path
End Sub

Private Sub File_Click()
Text1.Text = File.Path + "\" + File.FileName
End Sub

Private Sub Open_Click()
Text1.Visible = True
File.Visible = True
Directorio.Visible = True
unidad.Visible = True
Command1.Visible = True
End Sub

Private Sub unidad_Change()
Directorio.Path = unidad.Drive
End Sub
