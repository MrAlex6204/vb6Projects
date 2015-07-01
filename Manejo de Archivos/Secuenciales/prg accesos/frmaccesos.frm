VERSION 5.00
Begin VB.Form frmaccesos 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accesos"
   ClientHeight    =   4170
   ClientLeft      =   2955
   ClientTop       =   1890
   ClientWidth     =   5175
   Icon            =   "frmaccesos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "Introduca el Password"
      Top             =   2790
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Terminar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   1650
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   315
      TabIndex        =   8
      Top             =   2250
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accesos Al Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   585
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3225
   End
End
Attribute VB_Name = "frmaccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Archivo = FreeFile()
If VALIDAR = True Then
Open App.Path + "\Accesos.txt" For Append As #Archivo
Print #Archivo, UCase(Text1)
Print #Archivo, UCase(Text2)
Print #Archivo, UCase(Text3)
Close #Archivo
Limpia_Datos
MsgBox "usuario Agregado", vbInformation, "Verasoft Development"
'Procedimiento que permite agregar nuevos archivos con su paasword
'para que sean reconocidos por el sistema
Else
MsgBox "Favor de Llenar Todos los datos", vbCritical, "VeraSoft Deve.."
End If
End Sub
Sub Limpia_Datos()
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
'Procedimiento que limpia los datos
End Sub

Private Sub Command2_Click()
Dim xlogin As String
xlogin = InputBox("Dame el Login a Buscar", "Tectfiles")
If xlogin <> "" Then
 If BuscarLogin(xlogin) = True Then
 Text1.Text = Login
 Text2.Text = Nombre
 Text3.Text = Pwd
Else
 MsgBox "el usuario no existe"
End If
Close #Archivo
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Function VALIDAR() As Boolean
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
VALIDAR = False
Else
VALIDAR = True
End If
End Function


