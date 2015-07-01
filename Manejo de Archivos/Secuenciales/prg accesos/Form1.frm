VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3495
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2355
      TabIndex        =   2
      Top             =   3990
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   750
      TabIndex        =   1
      Top             =   3990
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1545
      TabIndex        =   0
      Top             =   3105
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   3510
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim log, nom, pass As String
If txtUserName.Text <> "" Then


Archivo = 15
Open App.Path + "\Accesos.txt" For Input As #Archivo

Do While EOF(Archivo)

       Input #Archivo, log
       Input #Archivo, nom
       Input #Archivo, pass
       
       List1.AddItem log
       List1.AddItem nom
       List1.AddItem pass
       
       
Loop

Close #Archivo


Else
MsgBox "Teclee el nom de usuario y contraseña", vbCritical, "Ouuch"
End If

End Sub

