Attribute VB_Name = "Mover"
Public Button As Integer, Shift As Integer, X As Single, Y As Single

Public lngReturnValue As Long
'Public Button As Integer, Shift As Integer, X As Single, Y As Single
Public Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Public Declare Sub ReleaseCapture Lib "user32" ()

      Public Const WM_NCLBUTTONDOWN = &HA1
      Public Const HTCAPTION = 2

Public mov As Long
Sub MoverForm(hWnd As Long)

Call ReleaseCapture
           
            lngReturnValue = SendMessage(hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
End Sub


'********************************************* _
TODO ESTO VA EN EL SUB DE MOUSEMOVE DEL FORM _
SINTAXIS PARA USAR ESTE MODULO _
EJEMPLO _

 'Dim lngReturnValue As Long _
 '
  '      If Button = 1 Then _
   '     Call ReleaseCapture _
    '    lngReturnValue = SendMessage(Form1.hWnd, WM_NCLBUTTONDOWN, _
     '   HTCAPTION, 0&)
      '  Mover.MoverForm
       ' End If
        

'*****************************************************************
' ESOT DEVE DE IR EN LA DECLARACION AL INICIO DEL PRG

'Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long


      'Private Declare Sub ReleaseCapture Lib "User32" ()
     
      

      'Const WM_NCLBUTTONDOWN = &HA1
      'Const HTCAPTION = 2

