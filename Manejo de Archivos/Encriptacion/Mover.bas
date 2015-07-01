Attribute VB_Name = "Mover"
Public lngReturnValue As Long
'Public Button As Integer, Shift As Integer, X As Single, Y As Single
Public Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Public Declare Sub ReleaseCapture Lib "User32" ()

    Public Const WM_NCLBUTTONDOWN = &HA1
    Public Const HTCAPTION = 2

Public mov As Long
Sub MoverForm()
mov = Form1.hWnd
Call ReleaseCapture
           
            lngReturnValue = SendMessage(mov, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
End Sub



