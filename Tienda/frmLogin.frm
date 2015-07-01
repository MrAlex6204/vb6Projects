VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   11010
   ClientLeft      =   2910
   ClientTop       =   3945
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6505.072
   ScaleMode       =   0  'User
   ScaleWidth      =   14309.53
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   1
      Top             =   4080
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   6840
      TabIndex        =   2
      Top             =   4560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   8040
      TabIndex        =   3
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cajero Num."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   360
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   4080
      Width           =   1590
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sItemData As String
   Dim strData As String 'Variable Que Sirvira Para Localizar Un Dato
   Dim strOutData As String
   Dim strConnect As String
   Public RS As New ADODB.Recordset
   Dim strPath As String
   Public NomCajero As String
   Public NumCajero As Integer
   





Private Sub cmdCancel_Click()
 Me.Hide
End Sub

 Sub cmdOK_Click()
AbrirBase
strData = "Cajero = '" & txtUserName.Text & "'"


    RS.Find strData
    If RS.EOF = True Then
    MsgBox "No Se Encuentra Cajero Registrado ", vbCritical, "No Se Encontro !!!!"
    CerrarBase
    Exit Sub
    End If
    
    NomCajero = RS.Fields("Nombre")
     
     
    



CerrarBase

    Me.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
 'Obtiene el Path de la Base de Datos
    strPath = App.Path & "\Tienda.mdb"
    
    'Variable Que Contine El Tipo de Coenccion que se va realizar
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      
      
      
End Sub
'Metodo Para Abrir La Base De Datos
Sub AbrirBase()
RS.Open "select * from Cajeros", strConnect, adOpenKeyset, adLockOptimistic
End Sub

'Metodo Para Cerrar la Base de Datos
Sub CerrarBase()
RS.Close
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AbrirBase
 strData = "Cajero = '" & txtUserName.Text & "'"


    RS.Find strData
    If RS.EOF = True Then
    MsgBox "No Se Encuentra Cajero Registrado ", vbCritical, "No Se Encontro !!!!"
    CerrarBase
    Exit Sub
    End If
    
    
     



CerrarBase
End If
End Sub
