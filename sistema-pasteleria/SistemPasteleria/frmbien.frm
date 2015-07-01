VERSION 5.00
Begin VB.Form frmbien 
   BackColor       =   &H8000000E&
   Caption         =   "Bienvenido"
   ClientHeight    =   3690
   ClientLeft      =   3885
   ClientTop       =   2565
   ClientWidth     =   4875
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4875
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3000
      Picture         =   "frmbien.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1320
      Picture         =   "frmbien.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblare 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Label lblcar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblnom 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Seguridad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3570
   End
End
Attribute VB_Name = "frmbien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRARSE AHORA", vbQuestion + vbYesNo, "Sistema de Seguridad")
If g = vbYes Then
copiarcampos
rsreg.Update
Else
rsreg.Cancel
End If
frmpresentacion.Show
frmbien.Hide
End Sub

Private Sub Form_Load()
activareg
llenarcampos
  frmbien.lblnom = frmreg.txtDNI(1)
    frmbien.lblcar = frmreg.txtDNI(2)
    frmbien.lblare = frmreg.txtDNI(3)
 'Format(Date, "long date") = lblfecha
 ' Format(Time, "long time") = lblhora




End Sub
Public Sub copiarcampos()
rsreg.Fields("nom_reg") = lblnom
rsreg.Fields("cargo") = lblcar
rsreg.Fields("area") = lblare
rsreg.Fields("fecha_ing") = Date
rsreg.Fields("hora_ing") = Time
rsreg.Fields("hora_sal") = Date
End Sub

Public Sub llenarcampos()
If rsreg.BOF Then Exit Sub
If rsreg.EOF Then Exit Sub
lblnom = rsreg.Fields("nom_reg")
lblcar = rsreg.Fields("cargo")
lblare = rsreg.Fields("area")
'rsreg.Fields("fecha_ing")
'rsreg.Fields("hora_ing") = Time
'rsreg.Fields("hora_sal") = Date





End Sub
