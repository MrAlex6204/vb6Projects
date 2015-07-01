VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form myform 
   Caption         =   "Cliente"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Text            =   "1007"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtIp 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
      Begin VB.Shape shpGo 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape shpWait 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   840
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shpError 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5520
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cerrar conexión"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Conectar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4680
      TabIndex        =   12
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección Ip"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de item"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "myform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################
'Author: S.S. Ahmed
'Email: ss_ahmed1@hotmail.com
'Date: Jul 21, 2001
'Note: This product is provided without any support
'#########################################################

Option Explicit

Private Sub cmdClose_Click()
Winsock1.Close
shpGo.Visible = False
shpWait.Visible = False
shpError.Visible = True
End Sub

Private Sub cmdConnect_Click()
Winsock1.RemoteHost = txtIp.Text  'Change this to your host ip
Winsock1.RemotePort = txtPort.Text
Winsock1.Connect
shpGo.Visible = True
txtItem.SetFocus
End Sub

Private Sub cmdSend_Click()
If Winsock1.State = sckConnected Then
    Winsock1.SendData txtItem.Text
    shpGo.Visible = True
    Label3.Caption = "Enviando datos"
Else
    shpGo.Visible = False
    shpWait.Visible = False
    shpError.Visible = True
    Label3.Caption = "No conectado al host"
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Winsock1.GetData sData, vbString
'Label1.Caption = sData
txtPrice.Text = sData
Label3.Caption = "Recibiendo datos"
shpGo.Visible = True
shpWait.Visible = False
shpError.Visible = False

End Sub

Private Sub Winsock1_SendComplete()
Label3.Caption = "Datos recibidos"
End Sub
