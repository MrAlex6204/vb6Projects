VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15240
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
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
      Left            =   6600
      TabIndex        =   1
      Top             =   5520
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   6600
      TabIndex        =   2
      Top             =   6000
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   7800
      TabIndex        =   3
      Top             =   6000
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   4320
      Picture         =   "frmLogin.frx":0000
      Top             =   960
      Width           =   7020
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
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   5520
      Width           =   1590
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

Sub cmdOK_Click()

If BuscarCajero(txtUserName) = True Then
CajeroOnline = Nombre
Open App.Path + "\" + CajeroOnline + ".Dat" For Output As #13
   Close #13
frmVentas.Show
Unload Me
Else
MsgBox "El Cajero No Existe", vbCritical, "VeraSoft Development"
End If

End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
'VALIDA SI EL  DATO INTRODUCIDO
'ES NUMERICO
If IsNumeric(txtUserName.Text) = False Then
Beep
txtUserName = Empty
Exit Sub
End If
If KeyAscii = 13 Then
Call cmdOK_Click
End If
End Sub
