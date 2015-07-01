Attribute VB_Name = "Mover"
Public lngReturnValue As Long
'Public Button As Integer, Shift As Integer, X As Single, Y As Single
Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "User32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2

Public mov As Long
Sub MoverForm()
mov = MDIForm1.hWnd
Call ReleaseCapture
           
            lngReturnValue = SendMessage(mov, WM_NCLBUTTONDOWN, _
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

