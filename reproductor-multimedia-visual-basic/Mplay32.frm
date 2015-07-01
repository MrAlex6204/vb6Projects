VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproductor Multimedia"
   ClientHeight    =   1095
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5385
   Icon            =   "Mplay32.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   3720
      Top             =   480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "(_)"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<<"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3240
      Top             =   480
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>|"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "||||"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "||"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2640
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu MnuFila 
      Caption         =   "&Fila"
      Begin VB.Menu MnuAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuCerrar 
         Caption         =   "&Cerrar"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu Nada 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MnuReproducir 
      Caption         =   "&Reproducción"
      Begin VB.Menu MnuReproduce 
         Caption         =   "&Reproducir"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuPausar 
         Caption         =   "&Pausar"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuDetener 
         Caption         =   "&Detener"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuResumir 
         Caption         =   "&Resumir"
         Shortcut        =   ^L
      End
      Begin VB.Menu Nada2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPanatafull 
         Caption         =   "&Pantalla Completa"
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuSonido 
         Caption         =   "&Sonido"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuConfigurar 
         Caption         =   "&Configurar"
         Shortcut        =   ^F
      End
      Begin VB.Menu Nada3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVelocidad 
         Caption         =   "&Velocidad"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu MnuAcercade 
         Caption         =   "&Acerca de..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim ShortName
Dim i As Long
Dim mssg As String * 255
Dim ResumeStat As String
Dim FFRR As String
 Private Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
      
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

      
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       
       GetShortName = Left(sShortPathName, lRetVal)
   End Function

Private Sub Command1_Click()
 i = mciSendString("play video1 from " & Slider1.Value, 0&, 0, 0)
DoEvents
End Sub

Private Sub Command2_Click()
 i = mciSendString("pause video1 ", 0&, 0, 0)
DoEvents
End Sub

Private Sub Command3_Click()
 i = mciSendString("stop video1 ", 0&, 0, 0)
DoEvents
End Sub

Private Sub Command4_Click()
 i = mciSendString("resume video1 ", 0&, 0, 0)
DoEvents
End Sub

Private Sub Command5_Click()
If MnuSonido.Caption = "Sonido" Then
       MnuSonido.Caption = "No Sonido"
       i = mciSendString("set video1 audio all off", mssg, 255, 0)
       Else
       i = mciSendString("set video1 audio all on", mssg, 255, 0)
       MnuSonido.Caption = "Sonido"
    End If

End Sub



Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all off", mssg, 255, 0)
i = mciSendString("status video1 mode", mssg, 255, 0)
FFRR = mssg
Timer2.Enabled = True
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all on", mssg, 255, 0)
Timer2.Enabled = False

Select Case Left$(FFRR, 4)
Case "stop"
 i = mciSendString("stop video1", 0&, 0, 0)
Case "play"
 i = mciSendString("play video1", 0&, 0, 0)
Case "paus"
 i = mciSendString("pause video1", 0&, 0, 0)
Case Else
End Select

End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all off", mssg, 255, 0)
i = mciSendString("status video1 mode", mssg, 255, 0)
FFRR = mssg
Timer3.Enabled = True
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all on", mssg, 255, 0)
Timer3.Enabled = False

Select Case Left$(FFRR, 4)
Case "stop"
 i = mciSendString("stop video1", 0&, 0, 0)
Case "play"
 i = mciSendString("play video1", 0&, 0, 0)
Case "paus"
 i = mciSendString("pause video1", 0&, 0, 0)
Case Else
End Select

End Sub

Private Sub Form_Resize()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29

End Sub

Private Sub Form_Unload(Cancel As Integer)
i = mciSendString("close all", 0&, 0, 0)
End
End Sub

Private Sub MnuAbrir_Click()
Dim MInfo As String
Screen.MousePointer = 11

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Archivos de sonido|*.wma;*.mp3;*.wav;*.aif;*.au|Archivo de Sonido|*.wav|Apple AIFF|*.aif|Windows Media Audio|*.wma|Sonido en formato mp3|*.mp3|Sonido en formato au|*.au|Secuencia Midi|*.mid;*.rmi|Pista de Audio En Cd |*.cda|Archivos de Video|*.wmv;*.mpa;*.mpe;*.mpg;*.mpeg;*.avi|Windows Media Video|*.wmv|Archivo de Pelicula(mpeg)|*.mpg;*.mpa;*.mpe;*.mpeg|Video para Windows|*.avi|Todos los ficheros (*.*)|*.*"
CommonDialog1.Flags = &H80000 Or &H1000
CommonDialog1.ShowOpen

'#####################################################
 ShortName = GetShortName(CommonDialog1.FileName)
'#####################################################

i = mciSendString("close all", 0&, 0, 0)
DoEvents
Get_Size GetShortName(CommonDialog1.FileName)
DoEvents
Me.Caption = "Reproductor Multimedia - " & CommonDialog1.FileTitle
App.Title = "Reproductor Multimedia - " & CommonDialog1.FileTitle
Screen.MousePointer = 0

Exit Sub

EH1:

Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
End Sub
Private Function Get_Size(ShortName As String)
Dim sReturn As String * 128
Dim lPos As Long
Dim lStart As Long

DoEvents
i = mciSendString("open " & ShortName & " Alias video1", 0&, 0, 0)
DoEvents
i = mciSendString("set video1 time format ms", 0&, 0, 0)
DoEvents
i = mciSendString("status video1 length", mssg, 255, 0)
DoEvents

Slider1.Max = Val(mssg)
DoEvents
 Timer1.Enabled = True
End Function

Private Sub MnuAcercade_Click()
Call ShellAbout(Me.hwnd, "Mplayer 1.0", "Copyright 2007, Dj_Dexter, es un clon del originar mplay32.exe ", Me.Icon)
End Sub

Private Sub MnuCerrar_Click()
i = mciSendString("close all", 0&, 0, 0)
Me.Caption = "Reproductor Multimedia"
App.Title = "Reproductor Multimedia"
End Sub

Private Sub MnuConfigurar_Click()
i = mciSendString("Configure video1", 0&, 0, 0)
End Sub

Private Sub MnuDetener_Click()
i = mciSendString("stop video1", mssg, 255, 0)

End Sub

Private Sub MnuPanatafull_Click()
 i = mciSendString("play video1 Fullscreen ", 0&, 0, 0)
End Sub

Private Sub MnuPausar_Click()
i = mciSendString("pause video1", mssg, 255, 0)
End Sub

Private Sub MnuReproduce_Click()
 i = mciSendString("play video1 from " & Slider1.Value, 0&, 0, 0)
DoEvents
End Sub

Private Sub MnuResumir_Click()
i = mciSendString("resume video1", mssg, 255, 0)
DoEvents
End Sub

Private Sub MnuSalir_Click()
i = mciSendString("close all", mssg, 255, 0)
DoEvents
End
End Sub

Private Sub MnuSonido_Click()
If MnuSonido.Caption = "Sonido" Then
       MnuSonido.Caption = "No Sonido"
       i = mciSendString("set video1 audio all off", mssg, 255, 0)
       Else
       i = mciSendString("set video1 audio all on", mssg, 255, 0)
       MnuSonido.Caption = "Sonido"
    End If

End Sub

Private Sub MnuVelocidad_Click()
On Error Resume Next
Dim yd As Integer
yd = InputBox("Pon un valor de 5 - 2200 para aumentar la velocidad", "Poner Velocidad", "1000")
 i = mciSendString("set video1 Speed " & yd & "", mssg, 255, 0)
If Err Then Beep

End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
i = mciSendString("status video1 mode", mssg, 255, 0)

If Left$(mssg, 7) = "playing" Then
  ResumeStat = "playing"
Else
  ResumeStat = ""
End If

i = mciSendString("pause video1", 0&, 0, 0)
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("seek video1 to " & Slider1.Value, 0&, 0, 0)

If ResumeStat = "playing" Then
  i = mciSendString("play video1", 0&, 0, 0)
End If
  
   
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim VidPos As String

i = mciSendString("status video1 position", mssg, 255, 0)
VidPos = Str(mssg)

Slider1.Value = VidPos
If Err Then Exit Sub

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
 i = mciSendString("stop video1", 0&, 0, 0)
 i = mciSendString("status video1 position", mssg, 255, 0)
 
 If mssg + 50 > Slider1.Max Then
     i = mciSendString("seek video1 to end", 0&, 0, 0)
 Else
    i = mciSendString("play video1 from " & mssg + 50, 0&, 0, 0)
 End If

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
 i = mciSendString("stop video1", 0&, 0, 0)
 i = mciSendString("status video1 position", mssg, 255, 0)
 
 If mssg - 50 <= 0 Then
    i = mciSendString("seek video1 to start", 0&, 0, 0)
    Slider1.Value = 0
 Else
    i = mciSendString("play video1 from " & mssg - 50, 0&, 0, 0)
 End If
End Sub

Private Sub Timer4_Timer()
Dim LTrackPosition, SegunI, MinutosI, SegundosI
Dim mssg As String * 255

  i = mciSendString("set video1 time format ms", 0&, 0, 0)
  i = mciSendString("status video1 Position", mssg, 255, 0)

LTrackPosition = mssg
SegunI = Val(LTrackPosition) \ 1000
MinutosI = SegunI \ 60
SegundosI = SegunI Mod 60
Label1.Caption = MinutosI & " min. " & SegundosI & " seg."

End Sub
