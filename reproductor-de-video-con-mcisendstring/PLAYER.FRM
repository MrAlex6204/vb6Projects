VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproductor de Video"
   ClientHeight    =   4230
   ClientLeft      =   270
   ClientTop       =   750
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "PLAYER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   5445
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
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
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MnuVideo 
      Caption         =   "&Video"
      Begin VB.Menu MnuReproducir 
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
      Begin VB.Menu MnuInicio 
         Caption         =   "&Inicio"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuResumir 
         Caption         =   "&Continuar"
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuReprocon 
         Caption         =   "&Reproducir continuado"
         Shortcut        =   ^U
      End
      Begin VB.Menu nada2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPantallacom 
         Caption         =   "&Pantalla completa"
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuVerFila 
         Caption         =   "&Ver Fila "
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuCuadroatra 
         Caption         =   "&Cuadro Atras"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuCuadroade 
         Caption         =   "&Cuadro Adelante"
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuSonido 
         Caption         =   "&Sonido"
         Shortcut        =   ^S
      End
      Begin VB.Menu nada3 
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
Dim path As String
Dim i As Long
Dim yd As Integer
Dim mssg As String * 255

Private Sub Form_Load()
Text1.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29

End Sub

Private Sub Form_Unload(Cancel As Integer)
i = mciSendString("close Video1", 0&, 0, 0)
End Sub

Private Sub MnuAbrir_Click()
Dim MInfo As String
Dim Corto As Long
Dim Last$, Todo$
Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hwnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "Archivos de Video" + Chr$(0) + "*.wmv;*.mpa;*.mpe;*.mpg;*.mpeg;*.avi" + Chr$(0) + "Windows Media Video" + Chr$(0) + "*.wmv" + Chr$(0) + "Archivo de Pelicula(mpeg)" + Chr$(0) + "*.mpg;*.mpa;*.mpe;*.mpeg" + Chr(0) + "Video para Windows" + Chr(0) + "*.avi" + Chr$(0) + "Todos los archivos" + Chr$(0) + "*.*" + Chr$(0)
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = CurDir
ofn.flags = &H4
ofn.lpstrTitle = "Abrir fila de Video"

Temp = ""
Temp = GetOpenFileName(ofn)

If Temp Then 'Si se ha elegido un fichero...
    Dim Fichero As String
    Fichero = Trim$(ofn.lpstrFile)
    Text1.Text = Fichero
End If
On Error GoTo EH1

i = mciSendString("close all", 0&, 0, 0)
'#####################################################
Last$ = Form1.hwnd & " Style " & &H40000000
PathCorto Fichero
DoEvents
Todo$ = "open " & path & " Alias Video1 parent " & Last$
DoEvents
i = mciSendString(Todo$, 0&, 0, 0)
i = mciSendString("put Video1 window at 1 1 361 280", 0&, 0, 0)
DoEvents
'#####################################################

Exit Sub

EH1:
Exit Sub
If Err = 32755 Then
MsgBox "Error: " & Error$
End If
Resume Next
Exit Sub


End Sub

Private Sub MnuAcercade_Click()
Call ShellAbout(Me.hwnd, "Reproductor de Video 1.0", "Copyright 2007, Dj_Dexter ", Me.Icon)
End Sub

Private Sub MnuCerrar_Click()
i = mciSendString("stop Video1", 0&, 0, 0)
i = mciSendString("close Video1", 0&, 0, 0)
Form1.Caption = "Reproductor de Video"
App.Title = "Reproductor de Video"

End Sub

Private Sub MnuCuadroade_Click()
i = mciSendString("Step Video1 By 1", mssg, 255, 0)
End Sub

Private Sub MnuCuadroatra_Click()
i = mciSendString("Step Video1 By -1", mssg, 255, 0)
End Sub

Private Sub MnuDetener_Click()
i = mciSendString("stop Video1", 0&, 0, 0)
End Sub

Private Sub MnuInicio_Click()
i = mciSendString("Seek Video1 To Start", mssg, 255, 0)
End Sub

Private Sub MnuPantallacom_Click()
i = mciSendString("Play Video1 FullScreen", 0&, 0, 0)
End Sub

Private Sub MnuPausar_Click()
i = mciSendString("pause Video1", 0&, 0, 0)
End Sub

Private Sub MnuReprocon_Click()
i = mciSendString("play Video1 repeat", 0&, 0, 0)

End Sub

Private Sub MnuReproducir_Click()
i = mciSendString("play Video1 from 0", 0&, 0, 0)

End Sub

Private Sub MnuResumir_Click()
Dim yd As Integer
i = mciSendString("Status Video1 Position", mssg, 255, 0)
yd = mssg
i = mciSendString("Play Video1 From " & yd & "", mssg, 255, 0)

End Sub

Private Sub MnuSalir_Click()
i = mciSendString("close Video1", 0&, 0, 0)
End
End Sub

Private Sub MnuSonido_Click()
If MnuSonido.Caption = "Sonido" Then
       MnuSonido.Caption = "No Sonido"
       i = mciSendString("set Video1 audio all off", mssg, 255, 0)
       Else
       i = mciSendString("set Video1 audio all on", mssg, 255, 0)
       MnuSonido.Caption = "Sonido"
    End If

End Sub

Private Sub MnuVelocidad_Click()
yd = InputBox("Pon un valor de 5 - 2200 para aumentar la velocidad", "Poner Velocidad", "1000")
i = mciSendString("set Video1 Speed " & yd & "", mssg, 255, 0)
If Err Then Beep

End Sub
Private Sub PathCorto(archivo As String)
Dim Temp As String * 250 'Buffer
Dim ret As Long
path = String(255, 0)
'Obtenemos el Path corto
ret = GetShortPathName(archivo, Temp, 164)
'Sacamos los nulos al path
path = Replace(Temp, Chr(0), "")

End Sub

Private Sub MnuVerFila_Click()
MsgBox Text1.Text, 64
End Sub
