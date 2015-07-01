VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskBroadcast 
      Left            =   5280
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   2014
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   1575
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call send 'llamamos a la funcion que nos permite enviar nuestro mensaje

End Sub
Private Sub Command2_Click()

End


End Sub

Private Sub Form_Load()
'Separamos el puerto 20145 para usarlo en nuestra

'aplicación.
Rem MI PUERTO DE ESCUCHA MI DATOS
wskBroadcast.Bind 2014

End Sub

Private Sub Text1_Change()
'Cuando el txtMensaje esté vacío, deshabilitar el botón

'de envío.

Command1.Enabled = (Len(Text1.Text) > 0)

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
Private Sub wskBroadcast_DataArrival(ByVal bytesTotal As Long)

Dim Datos As String 'Variable para guardar los datos


'Recibe los datos y los almacena en la variable
Rem RECIBE LOS DATOS POR EL PUERTO AL QUE LO ENVIARON
Rem EN ESTE CASO ES EL PUERTO POR DONDE ESCUCHO LOS DATOS YO
wskBroadcast.GetData Datos


'Si txtDatosRecibidos está vacío:

If Len(Text2.Text) = 0 Then

Text2.Text = wskBroadcast.RemoteHostIP & ">" & Datos

'de lo contrario insertar primero un salto de línea y

'luego los datos.

Else

Text2.Text = Text2.Text & vbCrLf & wskBroadcast.RemoteHostIP & ">" & Datos

End If

End Sub
Sub send()

On Error Resume Next 'Para ignorar error 126 en Win9X



'Es necesario establecer nuevamente el RemoteHost y

'el puerto, para asegurarse que los paquetes se lleguen

'a enviar a todos los destinatarios.

wskBroadcast.RemoteHost = "255.255.255.255"

Rem EL PUERTO AL QUE ENVIA LOS DATOS
wskBroadcast.RemotePort = 20145


wskBroadcast.SendData Text1.Text 'Envía los datos


Text1.Text = "" 'Limpia el txtMensaje

Text1.SetFocus 'Mueve el foco hacia txtMensaje

End Sub

