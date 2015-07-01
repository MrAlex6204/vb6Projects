VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ConeccionDatosGrales 
      Left            =   1440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   530
   End
   Begin VB.ListBox List3 
      Height          =   1620
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock ConeccionRecibe 
      Left            =   840
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   430
   End
   Begin MSWinsockLib.Winsock ConeccionEnvia 
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estatus:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Estatus As String

Private Sub Command1_Click()
On Error Resume Next
'Envia los Q Estoy Coenctados  a Todos Los que estan conectados
ConeccionEnvia.RemoteHost = "255.255.255.255"
' EL PUERTO AL QUE ENVIA LOS DATOS AL PUERTO 430 POR
'DONDE EL OTRO CONECTADO LO RECIBE
ConeccionEnvia.RemotePort = 430
ConeccionEnvia.SendData "CONECTADO"
End Sub


Private Sub ConeccionDatosGrales_DataArrival(ByVal bytesTotal As Long)
'Recibe en el Mismo Orden en
'el que se Enviaron los Datos
ConeccionDatosGrales.GetData RecibeIP

List1.AddItem ("<" + RecibeIP + ">")
List2.AddItem ("<" + RecibeIP + ">")
List3.AddItem ("<" + RecibeIP + ">")



End Sub


Private Sub ConeccionRecibe_DataArrival(ByVal bytesTotal As Long)
ConeccionRecibe.GetData Estatus
If Estatus = "CONECTADO" Then
Label1 = "Estatus:Conectado....*"
IP = ConeccionEnvia.LocalIP
'Envia Datos por la red al puerto 530
ConeccionRecibe.RemoteHost = "255.255.255.255"
ConeccionRecibe.RemotePort = 530
ConeccionRecibe.SendData IP 'Envia IP
ConeccionRecibe.SendData Usuario 'Envia Nom Usuario
ConeccionRecibe.SendData "Conectato" 'Envia Status
'-----------------------------------
End If
End Sub


Private Sub Form_Load()

'Puerto donde se recibe el Mensaje
ConeccionRecibe.Bind 430
ConeccionDatosGrales.Bind 530
End Sub
