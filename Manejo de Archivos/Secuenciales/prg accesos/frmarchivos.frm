VERSION 5.00
Begin VB.Form frmarchivos 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transacciones"
   ClientHeight    =   3630
   ClientLeft      =   2565
   ClientTop       =   2280
   ClientWidth     =   6150
   Icon            =   "frmarchivos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6150
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Terminar"
      Height          =   390
      Left            =   4080
      TabIndex        =   10
      Top             =   3000
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input(Leer)"
      Default         =   -1  'True
      Height          =   390
      Left            =   2760
      TabIndex        =   9
      Top             =   3000
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Append"
      Height          =   390
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Output"
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2430
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom. Emp"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   465
      TabIndex        =   6
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   645
      TabIndex        =   5
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Salario"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos de Empleados"
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
      Height          =   555
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5040
   End
End
Attribute VB_Name = "frmarchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim xnum As Integer
xnum = InputBox("Dame el numero de Empleado a Buscar", "Texfiles")
If Val(xnum) <> 0 Then
Archivo = FreeFile()
Open App.Path + "\empleados.txt" For Input As #Archivo
 Do While Not EOF(Archivo)
 
 Line Input #Archivo, Numero
 Line Input #Archivo, Nombre
 Line Input #Archivo, Salario
 
 If Numero = xnum Then
 Llena_Datos
 Exit Do
 Else
 MsgBox "el empleado no existe"
 Limpia_Datos
 End If
 Loop
Close #Archivo
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Archivo = FreeFile()
If VALIDAR = True Then
Open App.Path + "\empleados.txt" For Output As #Archivo
Print #Archivo, Text1.Text; "; "
Print #Archivo, Text2.Text
Print #Archivo, Text3.Text
Close #Archivo
Limpia_Datos
Command3.Enabled = False
Else
MsgBox "Por Favor Llene Todos los datos", vbCritical, "VeraSoft Deve.."
End If
End Sub

Private Sub Command4_Click()
Archivo = FreeFile()
Open App.Path + "\empleados.txt" For Append As #Archivo
'Open nombre for mode [Access access][Lock] as [#file number][Len=reelength]
Print #Archivo, Text1.Text
Print #Archivo, Text2.Text
Print #Archivo, Text3.Text
Close #Archivo
Limpia_Datos
End Sub

Private Sub Limpia_Datos()
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text1.SetFocus
End Sub
Private Sub Llena_Datos()
Text1.Text = Numero
Text2.Text = Nombre
Text3.Text = Salario
End Sub

Function VALIDAR() As Boolean
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
VALIDAR = False
Else
VALIDAR = True
End If
End Function
