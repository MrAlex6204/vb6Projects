Attribute VB_Name = "Globo"
Option Explicit
'Es una Sola variable Con Varias
'Variables Con Diferente Nombre
'o mas boen es una estructura de Datos
Private Type NOTIFYICONDATA
   Tamaño As Long
   hWnd As Long
   uID As Long
   IconStyle As Long
   Evento As Long
   Icono As Long
   Info As String * 128 ' variable con 128 caracteres
   dwState As Long
   dwStateMask As Long
   Texto As String * 256 ' variable con 256 caracteres
   Tiempo As Long
   Titulo As String * 64 ' variable con 64 caracteres
   GloboStyle As Long
End Type

'Es una variable tipo NOTIFYICONDATA
' O mas bien es una estructura de datos

Dim GloboPropiedadaes As NOTIFYICONDATA

 Const NOTIFYICON_VERSION = 3
 Const NOTIFYICON_OLDVERSION = 0
 
 Const NIM_ADD = &H0
 Const NIM_MODIFY = &H1
 Const NIM_DELETE = &H2

 Const NIM_SETFOCUS = &H3
 Const NIM_SETVERSION = &H4

 Const TipoMensaje = &H1
 Const TipoAlerta = &H2
 Const Icono = &H4

 Const NIF_STATE = &H8
 Const SinIcono = &H10

 Const NIS_HIDDEN = &H1
 Const NIS_SHAREDICON = &H2

 Const NIIF_NONE = &H0
 Const NIIF_WARNING = &H2
 Const NIIF_ERROR = &H3
 Const NIIF_INFO = &H1
 Const NIIF_GUID = &H4
'Nombre de Constantes de Eventos
'estos tipos de constantes tienen un valor

 Const WM_MOUSEMOVE = &H200
 Const WM_LBUTTONDOWN = &H201
 Const WM_LBUTTONUP = &H202
 Const WM_LBUTTONDBLCLK = &H203
 Const WM_RBUTTONDOWN = &H204
 Const WM_RBUTTONUP = &H205
 Const WM_RBUTTONDBLCLK = &H206
 
 'Shell_NotifyIcon es un funcion privada que es la declarada abajo
Private Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Sub AgregarIcono()
With GloboPropiedadaes
        .Tamaño = Len(GloboPropiedadaes)
        .hWnd = Form1.hWnd
        .uID = vbNull
        'Son Los Vario Tipos de Icononos del Globo
        .IconStyle = TipoAlerta Or SinIcono Or TipoMensaje Or Icono
        .Evento = WM_LBUTTONDBLCLK 'constante declarada que define como se oculta el Globo
        .Icono = Form1.Icon
        .Info = "Net Chat 1.0" & vbNullChar 'Muestra Informacion Rapida
        'al acercar el puntero sobre el icono en la  barra de Tareas
        ' y el vbNullChar elimina los Caracteres Nulos
        .dwState = 10
        .dwStateMask = 10
   End With
'NIM_ADD Es para agregar el  icono a la barra
   Shell_NotifyIcon NIM_ADD, GloboPropiedadaes
End Sub
Sub MostrarGlobo(MensajeEnTexto As String)
With GloboPropiedadaes
        .Tamaño = Len(GloboPropiedadaes)
        .hWnd = Form1.hWnd
        .uID = vbNull
        'Son Los Vario Tipos de Icononos del Globo
        .IconStyle = TipoAlerta Or SinIcono Or TipoMensaje Or Icono
        .Evento = WM_LBUTTONDBLCLK 'constante declarada que define como se oculta el Globo
        .Icono = Form1.Icon 'Muestra el icono del form en  la barra de tareas
        
        .Info = "Net Chat 1.0" & vbNullChar 'Muestra Informacion Rapida
        'al acercar el puntero sobre el icono en la  barra de Tareas
        ' y el vbNullChar elimina los Caracteres Nulos
        
        .dwState = 10
        .dwStateMask = 10
        .Texto = MensajeEnTexto & Chr(0) 'Texto del globo
        .Titulo = "Net Chat 1.0" & Chr(0) 'Titulo del globo
        .GloboStyle = TipoMensaje  'Selecionamos el tipo globo, de informacion en este caso)(NIIF_NONE, NIIF_INFO, NIIF_WARNING, NIIF_ERROR)
        .Tiempo = 1000 'Tiempo de espera  (millisec.)
   End With
  
  'Semanda a llamar la Funcion Shell_NotifyIcon
  
   Shell_NotifyIcon NIM_MODIFY, GloboPropiedadaes
   'Activamos el globo
End Sub


Sub Mensaje(Texto As String)

AgregarIcono
MostrarGlobo (Texto)

End Sub
