VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Servidor"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
      Begin VB.Label lblConnections 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblHostID 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblUsers 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Connections:"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIP 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblHostName 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "Host:"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "myform.frx":0000
      Left            =   120
      List            =   "myform.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   4080
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'############################################################
'Author: S.S. Ahmed
'Email: ss_Ahmed1@hotmail.com
'Date: Jul 21, 2001
'Note: This product is supplied without support of any kind.
'############################################################

Option Explicit
Dim iSockets As Integer
Dim sServerMsg As String
Dim sRequestID As String
   
Private Sub Form_Load()

    Form1.Show
    lblHostID.Caption = Socket(0).LocalHostName
    lblAddress.Caption = Socket(0).LocalIP
    Socket(0).LocalPort = 1007
    sServerMsg = "Listening to port: " & Socket(0).LocalPort
    List1.AddItem (sServerMsg)
    Socket(0).Listen
End Sub

Private Sub socket_Close(Index As Integer)
    sServerMsg = "Connection closed: " & Socket(Index).RemoteHostIP
    List1.AddItem (sServerMsg)
    Socket(Index).Close
    Unload Socket(Index)
    iSockets = iSockets - 1
    lblConnections.Caption = iSockets
    
End Sub

Private Sub socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    sServerMsg = "Connection request id " & requestID & " from " & Socket(Index).RemoteHostIP
  If Index = 0 Then
    List1.AddItem (sServerMsg)
    sRequestID = requestID
    iSockets = iSockets + 1
    lblConnections.Caption = iSockets
    Load Socket(iSockets)
    Socket(iSockets).LocalPort = 1007
    Socket(iSockets).Accept requestID
  End If

End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
   Dim sItemData As String
   Dim strData As String
   Dim strOutData As String
   Dim strConnect As String
   
        
    ' get data from client
    Socket(Index).GetData sItemData, vbString
    sServerMsg = "Received: " & sItemData & " from " & Socket(Index).RemoteHostIP & "(" & sRequestID & ")"
    List1.AddItem (sServerMsg)
   
    'strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\Prices.mdb;Persist Security Info=False"
    Dim strPath As String
     
    'Change the database path in the text file
     
    strPath = App.Path & "\prices.mdb"
    
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      
    Dim rs As New ADODB.Recordset
    
    ' Get clients request from database
    strData = "Item = '" & sItemData & "'"
    
    rs.Open "select * from prices", strConnect, adOpenKeyset, adLockOptimistic
    rs.Find strData
    strOutData = rs.Fields("Price")
    
    'send data to client
    sServerMsg = "Sending: " & strOutData & " to " & Socket(Index).RemoteHostIP
    List1.AddItem (sServerMsg)
    Socket(Index).SendData strOutData
    
End Sub


