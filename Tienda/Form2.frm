VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Alta Cajero"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNombre 
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
      IMEMode         =   3  'DISABLE
      Left            =   5865
      TabIndex        =   3
      Top             =   4230
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   7080
      TabIndex        =   2
      Top             =   4725
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   5760
      TabIndex        =   1
      Top             =   4725
      Width           =   1140
   End
   Begin VB.TextBox txtCajero 
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
      Left            =   5865
      TabIndex        =   0
      Top             =   3840
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cajero Num:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Index           =   1
      Left            =   3990
      TabIndex        =   5
      Top             =   3840
      Width           =   1770
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cajero Nom.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Index           =   0
      Left            =   3945
      TabIndex        =   4
      Top             =   4320
      Width           =   1860
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Dim SQL As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If txtCajero = Empty Or txtNombre = Empty Then
MsgBox "Porfavor Llene Todos Los Datos", vbInformation, "VeraSoft Develoment"
Exit Sub
End If

SQL = "INSERT INTO Cajeros (Cajero,Nombre) VALUES(" + txtCajero + ",'" + txtNombre + "')"
BDD.Execute SQL
BDD.Close
Unload Me
End Sub

Private Sub Form_Load()
'Hace la coneccion
Set BDD = OpenDatabase(App.Path & "\Tienda.mdb")

End Sub
