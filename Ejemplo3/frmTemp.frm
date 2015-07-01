VERSION 5.00
Begin VB.Form frmTemp 
   Caption         =   "Conversor de Temperaturas"
   ClientHeight    =   5625
   ClientLeft      =   3135
   ClientTop       =   2595
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Abrir"
      Height          =   1095
      Left            =   600
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Guardar"
      Height          =   1095
      Left            =   5040
      TabIndex        =   7
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtFahr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Text            =   "32"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtCent 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Text            =   "0"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.VScrollBar vsbTemp 
      Height          =   5055
      LargeChange     =   10
      Left            =   3840
      Max             =   -100
      Min             =   100
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblFahr 
      Caption         =   "Grados Farhenheit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lblCent 
      Caption         =   "Grados Centigrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub cmdOpen_Click()

End Sub

Private Sub cmdReset_Click()
vsbTemp.Value = 0
End Sub

Private Sub cmdSalir_Click()
End


End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub vsbTemp_Change()

txtCent.Text = vsbTemp.Value
txtFahr.Text = 32 + 1.8 * vsbTemp.Value

End Sub
