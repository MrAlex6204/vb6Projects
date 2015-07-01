VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Cahat Version 1.0"
   ClientHeight    =   5415
   ClientLeft      =   3510
   ClientTop       =   3015
   ClientWidth     =   7740
   Icon            =   "CLIENTE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7740
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "conectadto"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskBroadcast 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   2014
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4440
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
      Top             =   3720
      Width           =   4455
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
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
      Left            =   5775
      TabIndex        =   4
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   5040
      Picture         =   "CLIENTE.frx":0322
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2535
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

Private Sub menusobre_Click()
frmAbout.Show
End Sub

Private Sub mnuSalir_Click()
End
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

