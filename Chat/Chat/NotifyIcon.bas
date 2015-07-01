Attribute VB_Name = "Notify_Icon"
Option Explicit
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Dim nf_IconData As NOTIFYICONDATA

 Const NOTIFYICON_VERSION = 3
 Const NOTIFYICON_OLDVERSION = 0

 Const NIM_ADD = &H0
 Const NIM_MODIFY = &H1
 Const NIM_DELETE = &H2

 Const NIM_SETFOCUS = &H3
 Const NIM_SETVERSION = &H4

 Const NIF_MESSAGE = &H1
 Const NIF_ICON = &H2
 Const NIF_TIP = &H4

 Const NIF_STATE = &H8
 Const NIF_INFO = &H10

 Const NIS_HIDDEN = &H1
 Const NIS_SHAREDICON = &H2

 Const NIIF_NONE = &H0
 Const NIIF_WARNING = &H2
 Const NIIF_ERROR = &H3
 Const NIIF_INFO = &H1
 Const NIIF_GUID = &H4

 Const WM_MOUSEMOVE = &H200
 Const WM_LBUTTONDOWN = &H201
 Const WM_LBUTTONUP = &H202
 Const WM_LBUTTONDBLCLK = &H203
 Const WM_RBUTTONDOWN = &H204
 Const WM_RBUTTONUP = &H205
 Const WM_RBUTTONDBLCLK = &H206
 
Private Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Sub AgregarIcono()
With nf_IconData
        .cbSize = Len(nf_IconData)
        .hWnd = Form1.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Form1.icon
        .szTip = "Ballon-Tip Beispiel" & vbNullChar 'QuickInfo  Symbols & vbNullChar
        .dwState = 0
        .dwStateMask = 0
   End With

   Shell_NotifyIcon NIM_ADD, nf_IconData 'NIM_ADD Agregamos el icono a la barra
End Sub
Sub MostrarGlobo(Texto As String)
With nf_IconData
        .cbSize = Len(nf_IconData)
        .hWnd = Form1.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Form1.icon
        .szTip = "Balloon Maker" & vbNullChar 'QuickInfo  Symbols & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Texto & Chr(0) 'Texto del globo
        .szInfoTitle = "Chat" & Chr(0) 'Titulo del globo
        .dwInfoFlags = NIIF_INFO 'Selecionamos el tipo globo, de informacion en este caso)(NIIF_NONE, NIIF_INFO, NIIF_WARNING, NIIF_ERROR)
        .uTimeout = 1000 'Tiempo de espera  (millisec.)
   End With
 
   Shell_NotifyIcon NIM_MODIFY, nf_IconData 'Activamos el globo
End Sub
Sub QuitarIcono()
 Shell_NotifyIcon NIM_DELETE, nf_IconData 'NIM_DELETE Quitar el icono de la barra
End Sub
Sub Eventos(X As Single)
Dim lMsg As Long
   Dim sFilter As String
   lMsg = X / Screen.TwipsPerPixelX
   Select Case lMsg
   'you can play with other events as I did as per your use
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONUP
       Form2.PopupMenu Form2.MnuMostrar
      Case WM_LBUTTONDBLCLK
        Form2.Visible = True
        Form2.WindowState = vbNormal
      Case WM_RBUTTONDOWN
     
      Case WM_RBUTTONUP
       Form2.PopupMenu Form2.MnuMostrar
      'PopupMenu MnuIcono
      Case WM_RBUTTONDBLCLK
   End Select
End Sub


