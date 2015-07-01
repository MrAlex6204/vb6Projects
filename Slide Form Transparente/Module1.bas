Attribute VB_Name = "Module1"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''
' Código fuente en el módulo bas
''''''''''''''''''''''''''''''''''''''''''


' api para poner la ventana siempre visible
Private Declare Function SetWindowPos Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

' api para buscar el hwnd del Taskbar de windows
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

' api para obtener lasposición del TaskBar
Private Declare Function GetWindowRect Lib "User32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT) As Long

' Funciones api para la transparencia de la ventana
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal crey As Byte, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'constantes para la transparencia
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

' Constantes para poner la ventana alwaysOnTop
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

' acción para indicar si mostrar u ocultar el formulario
Enum EAccion
    MOSTRAR = 0
    OCULTAR = 1
End Enum

' Estructura que requiere el api GetWindowRect
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




' Función que muestra y oculta  el formulario en el systray
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SlideForm(FRM As Form, _
                     Direccion As EAccion, _
                     Optional LEVEL As Byte = 255, _
                     Optional Velocidad As Integer = 1)

Dim Posicion As Integer
Dim Tamaño As Integer
Dim hwnd As Long
Dim res As Long
Dim buffRECT As RECT

' Obtiene el hwnd de la barra de tareas de windows
hwnd& = FindWindow("Shell_TrayWnd", "")

If hwnd > 0 Then
    
    ' Obtiene las medidas para luego posicionar el formulario
    res = GetWindowRect(hwnd, buffRECT)
    
    If res > 0 Then
        ' Tamaño, es la Posición Top o Left  que tendrá el formulario al comienzo
        Tamaño = CStr(buffRECT.Bottom - buffRECT.Top) * 15
        
        ' Las posiciones, son las diferentes pos donde puede estar ubicada _
          la barra de tareas de windows ( abajo, arriba, izquierda derecha )
        
        If buffRECT.Left <= 0 And buffRECT.Top > 0 Then
            Posicion = 1
        End If
        
        If buffRECT.Left > 0 And buffRECT.Top <= 0 Then
            Posicion = 2
            Tamaño = (buffRECT.Right - buffRECT.Left) * 15
        End If
        
        If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right < 600 Then
            Posicion = 3
            Tamaño = buffRECT.Right * 15
        End If
        If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right > 600 Then
            Posicion = 4
        End If
    End If
Else
    Posicion = 1
End If

' Aplica la transparencia al formulario. El nivel o grado de transparencia _
    lo indica la variable LEVEL
Call SetWindowLong(FRM.hwnd, GWL_EXSTYLE, GetWindowLong(FRM.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(FRM.hwnd, 0, LEVEL, LWA_ALPHA)

' sentido de la animación de la ventana o formulario. Si es un 0 , es para mostrarlo _
  si se pasa el valor 1 es para ocultarlo
If Direccion = MOSTRAR Then
    FRM.Height = 0
    Select Case Posicion
        
        ' Posiciona el form
        Case 1
             FRM.Move Screen.Width - FRM.Width, Screen.Height - FRM.Height - Tamaño
        Case 2
             FRM.Move Screen.Width - FRM.Width - Tamaño, Screen.Height - FRM.Height
        Case 3
             FRM.Move Tamaño, Screen.Height - FRM.Height
        Case 4
             FRM.Move Screen.Width - FRM.Width, Tamaño
        End Select
    
    ' Pone el formulario siempre visible por encima de las demás _
      ventanas abiertas en windows ( modo alwaysOnTop)
    res = SetWindowPos(FRM.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or _
                       SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    
    
    ' animación de la ventana cuando se muestra
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do While FRM.Height <= 1515 ' la altura que se quiera
        DoEvents
        FRM.Height = FRM.Height + Velocidad
        If Not Posicion = 4 Then
            FRM.Top = FRM.Top - Velocidad
        End If
    Loop
Else
    ' Animación  Cuando se oculta y se descarga
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Do While FRM.Height >= 520
        DoEvents
        FRM.Height = FRM.Height - Velocidad
        If Not Posicion = 4 Then
            FRM.Top = FRM.Top + Velocidad
        End If
    Loop
    ' Descarga el formulario
    Unload FRM
End If
End Sub

