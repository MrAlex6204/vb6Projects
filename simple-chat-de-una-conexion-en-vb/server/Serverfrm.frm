VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerFrm 
   Caption         =   "Servidor"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton bntListen 
      Caption         =   "Poner a la escucha"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Tag             =   "Connect"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "123"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLog 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Escuchar en el puerto"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "ServerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bntListen_Click()
On Error GoTo errorSub

    With Winsock1
        .Close
        .LocalPort = txtPort
        .Listen
    End With

Exit Sub
errorSub:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub bntSend_Click()
On Error GoTo errorSub

    Winsock1.SendData txtSend
    
    txtLog = txtLog & "Servidor : " & txtSend & vbCrLf
    txtSend = ""

Exit Sub
errorSub:
MsgBox "Error : " & Err.Description
' cierra la conexión
Winsock1_Close
End Sub


Private Sub Winsock1_Close()
    ' Finaliza la conexión
    Winsock1.Close

    txtLog = txtLog & "*** Desconectado" & vbCrLf

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close ' close
    End If

    Winsock1.Accept requestID
    
    txtLog = "Cliente conectado. IP : " & _
              Winsock1.RemoteHostIP & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim dat As String

    Winsock1.GetData dat, vbString
    txtLog = txtLog & "Cliente : " & dat & vbCrLf

End Sub

' cuando se produce un error lo envía
''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
