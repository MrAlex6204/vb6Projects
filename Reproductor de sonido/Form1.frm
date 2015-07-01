VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Constantes para los flags

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  look for application specific association
Private Const SND_APPLICATION = &H80
'  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS = &H10000
'  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_ID = &H110000
'  play asynchronously
Private Const SND_ASYNC = &H1
  '  play synchronously (default)
Private Const SND_SYNC = &H0

'  name is a file name
Private Const SND_FILENAME = &H20000
'  loop the sound until next sndPlaySound
Private Const SND_LOOP = &H8
'  lpszSoundName points to a memory file
Private Const SND_MEMORY = &H4
'  silence not default, if sound not found
Private Const SND_NODEFAULT = &H2
 '  don't stop any currently playing sound
Private Const SND_NOSTOP = &H10
 '  don't wait if the driver is busy
Private Const SND_NOWAIT = &H2000
 '  purge non-static events for task
Private Const SND_PURGE = &H40
 '  name is a resource name or atom
Private Const SND_RESOURCE = &H40004

' Declaración del api PlaySound
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long

Private Sub Command1_Click()
    With CommonDialog1
        .DialogTitle = " Seleccionar archivo de audio"
        ' Filtra los Archivos con extensión wav
        .Filter = "Archivos wav|*.wav"
        
        ' Abre el diálogo
        .ShowOpen
        
        If .FileName = vbNullString Then
            Exit Sub
        Else
            Text1.Text = .FileName
        End If
    End With
    
End Sub

' Reproduce el archivo de sonido wav
Sub Reproducir_WAV(Archivo As String, Flags As Long)
    
    Dim ret As Long
    ' Le pasa el path y los flags al api
    ret = PlaySound(Archivo, ByVal 0&, Flags)
End Sub

' Botón para reproducir el sonido
Private Sub Command2_Click()
    
    Call Reproducir_WAV(Text1.Text, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT)
End Sub

Private Sub Form_Load()
    
    Command1.Caption = "Abrir archivo "
    Command2.Caption = "Reproducir"
    Me.Caption = "PlaySound"
End Sub
 

