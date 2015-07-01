VERSION 5.00
Begin VB.Form frmreg 
   BackColor       =   &H8000000B&
   Caption         =   "Registro"
   ClientHeight    =   3375
   ClientLeft      =   3915
   ClientTop       =   2970
   ClientWidth     =   4230
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4230
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "Ingrese su Password"
      Top             =   2040
      Width           =   2715
   End
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   2715
   End
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   2715
   End
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1200
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Ingrese su Password"
      Top             =   960
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESAR"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   690
   End
End
Attribute VB_Name = "frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If txtDNI(0) = "" Then
MsgBox "INGRESE SU PASSWORD", vbCritical, "SISTEMA DE SEGURIDAD"
End If

frmpresentacion.Show
frmreg.Hide

End Sub


Private Sub Form_Load()
conectar
activareg
llenarcampos




End Sub

Private Sub txtDNI_Change(Index As Integer)
txtDNI(0).SetFocus
Select Case Index
        Case 0
            If Len(txtDNI(0)) = 8 Then
                
                activausu
                
                rsusu.MoveFirst
                rsusu.Find "DNI_usu ='" + Trim(txtDNI(0)) + "'"
                If Not rsusu.EOF Then
                    txtDNI(1) = rsusu.Fields("nom_usu")
                    txtDNI(2) = rsusu.Fields("nom_car")
                    txtDNI(3) = rsusu.Fields("cod_are")
                    'frameCabecera.Enabled = True
                End If
                If Not txtDNI(1) = "" Then
                    txtDNI(1).Enabled = False
                    txtDNI(2).Enabled = False
                    txtDNI(3).Enabled = False
                  Else
                    txtDNI(1).Enabled = True
                    txtDNI(2).Enabled = True
                    txtDNI(3).Enabled = True
                    
                   txtDNI(0).SetFocus
                'frameCabecera.Enabled = True
                End If
            Else
                txtDNI(1) = ""
                txtDNI(2) = ""
                txtDNI(3) = ""
                
            End If
    End Select

End Sub


Public Sub copiarcampos()

rsreg.Fields("nom_reg") = txtDNI(1)
rsreg.Fields("cargo") = txtDNI(2)
rsreg.Fields("area") = txtDNI(3)
rsreg.Fields("fecha_ing") = Date
rsreg.Fields("hora_ing") = Time
rsreg.Fields("hora_sal") = Date
End Sub

Public Sub llenarcampos()
If rsreg.BOF Then Exit Sub
If rsreg.EOF Then Exit Sub
End Sub
