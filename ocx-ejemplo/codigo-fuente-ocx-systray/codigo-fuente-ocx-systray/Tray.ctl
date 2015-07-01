VERSION 5.00
Begin VB.UserControl Tray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "Tray.ctx":0000
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Función Api Shell_NotifyIcon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal _
                                              dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


'*******************************************
'Eventos
'*******************************************

Public Event MouseDown(Button As Integer)
Public Event MouseUP(Button As Integer)
Public Event MouseMove()
Public Event DblClick(Button As Integer)



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



'Coloca el systray
'************************
Public Sub PonerSystray()

  'Tamaño de la estructura systray
  sysTray.cbSize = Len(sysTray)
  'Establecemos el Hwnd, en este caso del formulario
  sysTray.hwnd = UserControl.hwnd
  sysTray.uId = vbNull
  'Flags
  sysTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  'Establecemos el mensaje callback
  sysTray.ucallbackMessage = WM_MOUSEMOVE
  'establecemos el icono, en este caso el que tiene el control Image1
  sysTray.hIcon = Image1.Picture
  'Establecemos el tooltiptext
  sysTray.szTip = m_ToolTiptext & vbNullChar
  'Ponemos el icono en el systray
  Shell_NotifyIcon NIM_ADD, sysTray

End Sub



'Remueve el systray
'************************
Public Sub RemoverSystray()
Shell_NotifyIcon NIM_DELETE, sysTray
End Sub


'Propiedad ToolTiptext
'*********************
Public Property Get ToolTipText() As String
ToolTipText = m_ToolTiptext
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
  m_ToolTiptext = vNewValue
  sysTray.szTip = m_ToolTiptext & vbNullChar
  Shell_NotifyIcon NIM_MODIFY, sysTray
  PropertyChanged ("ToolTipText")
End Property


'Propiedad IconPicture
'*********************
Public Property Get IconPicture() As Picture
    Set IconPicture = Image1.Picture
End Property

Public Property Set IconPicture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    sysTray.hIcon = Image1.Picture
    Shell_NotifyIcon NIM_MODIFY, sysTray
    PropertyChanged "IconPicture"
    UserControl_Resize
End Property

Property Let IconPicture(ByVal New_Picture As Picture)
  Set Image1.Picture = New_Picture
  PropertyChanged "IconPicture"
  UserControl_Resize
End Property


'Eventos del UserControl
'****************************
Private Sub UserControl_Initialize()
Image1.Top = 0
Image1.Left = 0

End Sub

Private Sub UserControl_Resize()
    Static flag As Boolean
    
    If flag Then Exit Sub
    
    flag = True
    
    With Image1
         Height = .Height
         Width = .Width
    
    End With
    flag = False
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mensaje As Long

On Local Error Resume Next

If (ScaleMode = vbPixels) Then
    mensaje = X
Else
    mensaje = X / Screen.TwipsPerPixelX
End If

RaiseEvent MouseMove

Select Case mensaje
       
       'Dobleclick boton izquierdo
       Case WM_LBUTTONDBLCLK
            RaiseEvent DblClick(vbLeftButton)
       'Dobleclick boton derecho
       Case WM_RBUTTONDBLCLK
            RaiseEvent DblClick(vbRightButton)
       'Botón Arriba Derecho
       Case WM_RBUTTONUP
            RaiseEvent MouseUP(vbRightButton)
       'Botón Arriba Izquierdo
       Case WM_LBUTTONUP
            RaiseEvent MouseUP(vbLeftButton)
       'Botón Derecho abajo
       Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)
       'Botón izquierdo abajo
       Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton)
End Select

End Sub



'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Local Error Resume Next
    m_ToolTiptext = PropBag.ReadProperty("ToolTipText", "")
    Set Image1.Picture = PropBag.ReadProperty("IconPicture", Nothing)
End Sub


'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Local Error Resume Next
    Call PropBag.WriteProperty("ToolTipText", m_ToolTiptext, "")
    Call PropBag.WriteProperty("IconPicture", Image1.Picture, Nothing)
End Sub



