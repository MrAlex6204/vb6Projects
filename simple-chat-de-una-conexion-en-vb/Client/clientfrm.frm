VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientFrm 
   Caption         =   "Cliente"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton bntConnect 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Tag             =   "Connect"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "123"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtLog 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Puerto"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP remota"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "ClientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bntConnect_Click()
On Error GoTo ErrSub

    With Winsock1
        .Close
        .RemoteHost = txtIP
        .RemotePort = txtPort
        .Connect
    End With
Exit Sub
ErrSub:
MsgBox "Error : " & Err.Description, vbCritical
End Sub


Private Sub bntSend_Click()
On Error GoTo ErrSub


    Winsock1.SendData txtSend

    txtLog = txtLog & "Cliente : " & txtSend & vbCrLf
    txtSend = ""

Exit Sub
ErrSub:
MsgBox "Error : " & Err.Description
Winsock1_Close ' cierra la conexión
End Sub

Private Sub Winsock1_Close()

    Winsock1.Close  'Cierra la conexión
    txtLog = txtLog & "*** Desconectado" & vbCrLf

End Sub

Private Sub Winsock1_Connect()

txtLog = "Conectado a " & Winsock1.RemoteHostIP & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim dat As String
    
    Winsock1.GetData dat, vbString
    txtLog = txtLog & "Servidor : " & dat & vbCrLf

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, _
                           Description As String, _
                           ByVal Scode As Long, _
                           ByVal Source As String, _
                           ByVal HelpFile As String, _
                           ByVal HelpContext As Long, _
                           CancelDisplay As Boolean)

    txtLog = txtLog & "*** Error : " & Description & vbCrLf

    Winsock1_Close
End Sub
