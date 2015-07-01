Attribute VB_Name = "Module1"
Option Explicit

'Declare needed functions from Windows API
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    y, ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

'Always on Top constants
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


'Keyboard related Constants and Structs
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

'Keyboard related variables
'Dim p As KBDLLHOOKSTRUCT
Public IdKeyBoard As Long

'Mouse related Constants and Structs
Public Const WH_MOUSE_LL = 14


'Mouse related variables
'Dim p2 As MSLLHOOKSTRUCT
Public IdMouse As Long

'función que desactiva el teclado
'''''''''''''''''''''''''''''''''
Public Function WinProcKeyBoard(ByVal nCode As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long

    WinProcKeyBoard = -1

End Function

'función que desactiva el Mouse
'''''''''''''''''''''''''''''''
Public Function WinProcMouse(ByVal nCode As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long


WinProcMouse = -1
            
        
End Function
          

