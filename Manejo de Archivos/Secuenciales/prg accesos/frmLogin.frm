VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   3960
   ClientTop       =   4080
   ClientWidth     =   4815
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   4521.024
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()



If UCase(txtUserName.Text) <> "" Then


Archivo = 50
Open App.Path + "\Accesos.txt" For Input As #Archivo

Do While Not EOF(Archivo)

      Line Input #Archivo, Login
      Line Input #Archivo, Nombre
      Line Input #Archivo, Pwd
       
       If UCase(txtUserName) = Login Then
      
            If UCase(txtPassword) = Pwd Then
            frmmenu.Show
            Unload Me
            Close #Archivo
            Exit Sub
            Else
            MsgBox "Contraseña Incorrecta", vbCritical, "Ouuch"
            End If
        
        
       Else
       MsgBox "Usuario Incorrecto", vbCritical, "Ouuch"
      End If
Loop

Close #Archivo


Else
MsgBox "Teclee el nom de usuario y contraseña", vbCritical, "Ouuch"
End If

End Sub

