VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Cahat Version 1.0"
   ClientHeight    =   7710
   ClientLeft      =   2715
   ClientTop       =   1605
   ClientWidth     =   12600
   Icon            =   "frmNetChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12600
   Begin VB.ListBox List2 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmNetChat.frx":0322
      Left            =   10320
      List            =   "frmNetChat.frx":0324
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmNetChat.frx":0326
      Left            =   6360
      List            =   "frmNetChat.frx":0328
      TabIndex        =   9
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Timer Conectado 
      Interval        =   10
      Left            =   4920
      Top             =   5760
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   7335
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Conectado....."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskBroadcast 
      Left            =   1560
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   2014
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock WinsockUser 
      Left            =   2160
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   80
   End
   Begin MSWinsockLib.Winsock WinsockConectados 
      Left            =   2760
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   90
   End
   Begin MSWinsockLib.Winsock WinsockEscribiendo 
      Left            =   3480
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   100
   End
   Begin MSWinsockLib.Winsock IPSend 
      Left            =   4080
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   8080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "Conectados:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "Net Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   285
      TabIndex        =   6
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "Net Chat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   5
      Top             =   6480
      Width           =   1290
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   6840
      Width           =   1275
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2895
      Left            =   6480
      Picture         =   "frmNetChat.frx":032A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3615
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu menuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu menusobre 
         Caption         =   "&Sobre NetCchat"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public User As String
Dim IP As String

Private Sub Command1_Click()
Call send 'llamamos a la funcion que nos permite enviar nuestro mensaje

End Sub
Private Sub Command2_Click()

End


End Sub

Private Sub Conectado_Timer()
Call Sendconect 'Avisa a Todos que me Conecte

Call SendUser

wskBroadcast.RemoteHost = "255.255.255.255"
' EL PUERTO AL QUE ENVIA LOS DATOS 20145
wskBroadcast.RemotePort = 20145
wskBroadcast.SendData ".....RECIEN CONECTADO" 'Envía los datOS

StatusBar1.SimpleText = "          " 'se borra el texto de statusbar una vez despues de que
'se enviaron los datos
StatusBar1.Refresh


Conectado.Enabled = False
End Sub

Private Sub Form_Load()

Label2.Caption = frmLogin.txtUserName.Text
Unload frmLogin
'Separamos el puerto 20145 para usarlo en nuestra
'aplicación.
Rem MI PUERTO DE ESCUCHA MI DATOS
wskBroadcast.Bind 20145 'Puerto a donde llega el Mensaje
WinsockUser.Bind 80  'Puerto a donde llega el Nom. de Usuario
WinsockConectados.Bind 90 'Puerto donde LLega el Los Q estan Conectados
WinsockEscribiendo.Bind 91  'puerto donde envia el estatus de quien esta escribiendo



End Sub

Private Sub IPSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub List1_Click()
Label3 = List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()

End Sub

Private Sub menusobre_Click()
frmAbout.Show
End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
'Envia los Q Estoy Coenctados  a Todos Los que Me escuchen
WinsockUser.RemoteHost = "255.255.255.255"
Rem EL PUERTO AL QUE ENVIA LOS DATOS 90
WinsockUser.RemotePort = 90
WinsockUser.SendData "*Desconectado: " + Label2.Caption + "*" 'Envía los datos

Rem EL PUERTO AL QUE ENVIA LOS DATOS 80
WinsockUser.RemoteHost = "255.255.255.255"
WinsockUser.RemotePort = 80
WinsockUser.SendData "*Desconectado:" + User

Rem Cierra todos los Winsocks que tenga Abiertos
wskBroadcast.Close
WinsockUser.Close
WinsockConectados.Close
WinsockEscribiendo.Close
Unload Me
End
End Sub

Private Sub Text1_Change()
On Error Resume Next
'Cuando el txtMensaje esté vacío, deshabilitar el botón
'de envío.

Command1.Enabled = (Len(Text1.Text) > 0)

If (Len(Text1.Text) = 0) Then

'Envia a todos por la red que estoy escribiendo
WinsockEscribiendo.RemoteHost = "255.255.255.255"
'envia a todos el mensaje por el puerto 91 de la red
Rem EL PUERTO AL QUE ENVIA LOS DATOS 91
WinsockEscribiendo.RemotePort = 91
WinsockEscribiendo.SendData "          "


Else
Call Status

End If


End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then

Call send 'llamamos a la funcion que nos permite enviar nuestro mensaje

End If

End Sub

Private Sub Text2_Change()
'Mostrar siempre la última línea del TextBox.
Text2.SelStart = Len(Text2.Text)


End Sub



Private Sub txtConectados_Change()

End Sub

Private Sub WinsockConectados_DataArrival(ByVal bytesTotal As Long)

Dim conectados  As String
WinsockConectados.GetData conectados



List1.AddItem ("<" + conectados + ">")

StatusBar1.SimpleText = "Se Acaba de Conectar el Usuario:" + conectados
Globo.Mensaje (conectados)
Load frmSplash1
frmSplash1.Show

Beep
End Sub

Private Sub WinsockEscribiendo_DataArrival(ByVal bytesTotal As Long)
Dim Status As String
WinsockEscribiendo.GetData Status
StatusBar1.SimpleText = Status

End Sub

Private Sub WinsockUser_DataArrival(ByVal bytesTotal As Long)



'Recibe los datos y los almacena en la variable
Rem RECIBE LOS DATOS POR EL PUERTO AL QUE LO ENVIARON
Rem EN ESTE CASO ES EL PUERTO POR DONDE ESCUCHO LOS DATOS YO
WinsockUser.GetData User








End Sub

Private Sub wskBroadcast_DataArrival(ByVal bytesTotal As Long)
'NOTA
'CAPTURA TODOS LOS DATOS ENVIADOS POR LA RED
'POR EL PUERTO QUE TIENE ASIGNADO QUE ES 20145
'POR ESTE PUERTO SE ENVIA LOS MENSAJES
Dim Datos As String 'Variable para guardar los datos
'Recibe los datos y los almacena en la variable
Rem RECIBE LOS DATOS POR EL PUERTO AL QUE LO ENVIARON
Rem EN ESTE CASO ES EL PUERTO POR DONDE ESCUCHO LOS DATOS YO
wskBroadcast.GetData Datos


If Datos = ".....RECIEN CONECTADO" Then
List1.Clear
Call SendUser
Else
    'Si txtDatosRecibidos está vacío:
    If Len(Text2.Text) = 0 Then
    Text2.Text = "<" & User & ">" & Datos
    'de lo contrario insertar primero un salto de línea y
    'luego los datos.
    Else
    Text2.Text = Text2.Text & vbCrLf & "<" & User & ">" & Datos
    End If

End If
End Sub
Sub send()
On Error Resume Next 'Para ignorar error 126 en Win9X

Call SendUser 'Envia los datos de la funcion SendUser


'Es necesario establecer nuevamente el RemoteHost y
'el puerto, para asegurarse que los paquetes se lleguen
'a enviar a todos los destinatarios.
wskBroadcast.RemoteHost = "255.255.255.255"
' EL PUERTO AL QUE ENVIA LOS DATOS 20145
wskBroadcast.RemotePort = 20145
wskBroadcast.SendData Text1.Text 'Envía los datos


Text1.Text = "" 'Limpia el txtMensaje

Text1.SetFocus 'Mueve el foco hacia txtMensaje
StatusBar1.SimpleText = "          " 'se borra el texto de statusbar una vez despues de que
StatusBar1.Refresh
'se enviaron los datos


End Sub
Sub SendUser()

Rem Esta Funcion Sirve para Enviar
Rem el Nombre de Usuario
Rem al momento de enviar el texto
Rem y es para saber quien envio el mensaje
Rem o para ver de quien es el mensaje q se envio
On Error Resume Next
'Envia Nombre de Usuario
' al puerto 80
WinsockUser.RemoteHost = "255.255.255.255"
Rem EL PUERTO AL QUE ENVIA LOS DATOS 80
WinsockUser.RemotePort = 80
WinsockUser.SendData Label2.Caption  'Envía los datos label2

End Sub
Sub Sendconect()
On Error Resume Next
'Envia los Q Estoy Coenctados  a Todos Los que estan conectados
WinsockUser.RemoteHost = "255.255.255.255"
Rem EL PUERTO AL QUE ENVIA LOS DATOS 90
WinsockUser.RemotePort = 90
' envia el siguiente texto por la red
WinsockUser.SendData Label2.Caption  'Envía los datos

End Sub
Sub Status()
On Error Resume Next
'Envia a todos por la red que estoy escribiendo
WinsockEscribiendo.RemoteHost = "255.255.255.255"
'envia a todos el mensaje por el puerto 91 de la red
Rem EL PUERTO AL QUE ENVIA LOS DATOS 91
WinsockEscribiendo.RemotePort = 91
WinsockEscribiendo.SendData "<Esta Escribiendo: " + Label2.Caption + ">" 'Envía los datos

End Sub
Sub SendIP()
IP = IPSend.LocalIP
On Error Resume Next
'Envia los Q Estoy Coenctados  a Todos Los que estan conectados
IPSend.RemoteHost = "255.255.255.255"
Rem EL PUERTO AL QUE ENVIA LOS DATOS 8080
IPSend.RemotePort = 8080
WinsockUser.SendData IP  'Envía la ip

End Sub
