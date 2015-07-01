VERSION 5.00
Begin VB.UserControl SysTray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Función Api Shell_NotifyIcon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal _
                                              dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


'*******************************************
'Estuctura
'*******************************************

'Estructura NOTIFYICONDATA
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

'*******************************************
'Constantes
'*******************************************
'Para las acciones de Shell_NotifyIcon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Para los botones y el mouse (mensajes)
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

'*******************************************
'Variables locales
'*******************************************

'Variables para las propiedades
Private m_ToolTiptext As String

'*******************************************
'otras Variables
'*******************************************

'variable para la estructura NOTIFYICONDATA
Dim sysTray As NOTIFYICONDATA

Private Sub UserControl_Initialize()

End Sub
