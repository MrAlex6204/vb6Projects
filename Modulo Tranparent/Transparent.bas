Attribute VB_Name = "Transparent"
'*********************************************************************** _
formato de como de debe mandar a llamar la funcion del modulo _
Transparent.Aplicar_Transparencia frmSplash.hwnd, 215

'Declaracion del API SetLayerdWindowsAttributes que etablece la _
transparencia _
***********************************************************************
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crkey As Long, ByVal bAlph As Byte _
, ByVal dwFlas As Long) As Long
'**********************************************************************

'Recupera el Estilo de la Ventana _
***********************************************************************
Private Declare Function GetWindowsLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long
'**********************************************************************


'Decalracion del API SetWindowLong necesaria para aplicar un estilo _
al form antes de usar el API SetLayeredWindowAttributes _
***********************************************************************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'**********************************************************************

'**********************************************************************
Private Const GWL_EXTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'**********************************************************************

'Funcion para saber si el formulario ya  estransparente _
se le pasa el Hwnd del formulario en cuestion

 Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
 On Error Resume Next
 Dim Msg As Long
 If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  Msg = GetWindowsLong(hwnd, GWL_EXTYLE)
  Is_Transparent = True
  Else
  End If
  
  If Err Then
  Is_Transparent = False
  End If
 End Function

'funcion que aplica la transparencia ,se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer)
Dim Msg As Long
'On Error Resume Next
If Valor < 0 Or Valor > 255 Then
Aplicar_Transparencia = 1
Else
Msg = GetWindowsLong(hwnd, GWL_EXTYLE)
Msg = Msg Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXTYLE, Msg
'Establece la transparencia
SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA

Aplicar_Transparencia = 0
End If

If Err Then
Aplicar_Transparencia = 2

End If

End Function

