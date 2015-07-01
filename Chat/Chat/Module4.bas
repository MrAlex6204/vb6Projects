Attribute VB_Name = "Slide_Form"
Option Explicit
'---------form on top
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTTOP = 1
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
'----------Buscar la posicion del la barra de herramientas
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
'-------Transparencia
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public YaAbierto As Boolean

Public Sub SlideForm(FRM As Form, Direccion As Long)
Dim Posicion As Integer
Dim Tamaño As Integer
Dim hWnd As Long
Dim res As Long
Dim buffRECT As RECT
Dim FormAbierto As Integer
If YaAbierto Then FormAbierto = 2500


hWnd& = FindWindow("Shell_TrayWnd", "")
If hWnd > 0 Then
res = GetWindowRect(hWnd, buffRECT)
If res > 0 Then
Tamaño = CStr(buffRECT.Bottom - buffRECT.Top) * 15
If buffRECT.Left <= 0 And buffRECT.Top > 0 Then Posicion = 1
If buffRECT.Left > 0 And buffRECT.Top <= 0 Then Posicion = 2: Tamaño = (buffRECT.Right - buffRECT.Left) * 15
If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right < 600 Then Posicion = 3: Tamaño = buffRECT.Right * 15
If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right > 600 Then Posicion = 4
End If
Else
Posicion = 1
End If

If Direccion = 0 Then
res = SetWindowPos(FRM.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
FRM.Image1(0).Picture = Form1.Comando(2).Picture
FRM.Command1.Picture = Form1.Picture
FRM.Command2.Picture = Form1.Picture
Call SetWindowLong(FRM.hWnd, GWL_EXSTYLE, GetWindowLong(FRM.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(FRM.hWnd, 0, 200, LWA_ALPHA)


Dim x           As Integer
Dim y           As Integer
y = 0
While y < 2500
    x = 0
    While x < FRM.Width
       FRM.PaintPicture Form1.Picture, x, y
       FRM.PaintPicture Form1.Borde(0).Picture, x, 0
        x = x + Form1.Picture.Width \ 2
    Wend
        y = y + Form1.Picture.Height \ 2
    Wend
sndplaysound "C:\WINDOWS\Media\chimes.wav", SND_NODEFAULT + SND_ASYNC
FRM.Show
FRM.Height = 0
Select Case Posicion
Case 1
FRM.Move Screen.Width - FRM.Width, Screen.Height - FRM.Height - Tamaño - FormAbierto
Case 2
FRM.Move Screen.Width - FRM.Width - Tamaño, Screen.Height - FRM.Height - FormAbierto
Case 3
FRM.Move Tamaño, Screen.Height - FRM.Height - FormAbierto
Case 4
FRM.Move Screen.Width - FRM.Width, Tamaño + FormAbierto
End Select

Do Until FRM.Height = 2500 ' la altura que se quiera
DoEvents
FRM.Height = FRM.Height + 1
If Not Posicion = 4 Then FRM.Top = FRM.Top - 1

Loop



YaAbierto = True
Else
Do Until FRM.Height = 520
DoEvents
FRM.Height = FRM.Height - 1
If Not Posicion = 4 Then FRM.Top = FRM.Top + 1
Loop
Unload FRM
YaAbierto = False
End If
End Sub


