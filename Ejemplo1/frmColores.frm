VERSION 5.00
Begin VB.Form frmColores 
   Caption         =   "Colores"
   ClientHeight    =   6330
   ClientLeft      =   4350
   ClientTop       =   3795
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7245
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtCaja 
      Height          =   1455
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Frame fraPosición 
      Caption         =   "Posicíon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
      Begin VB.OptionButton optAbajo 
         Caption         =   "Abajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optArriba 
         Caption         =   "Arriba"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraColores 
      Caption         =   "Colores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton optAmarillo 
         Caption         =   "Amarillo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton optVerde 
         Caption         =   "Verde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -960
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optRojo 
         Caption         =   "Rojo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optAzul 
         Caption         =   "Azul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmColores.Hide

End Sub

Private Sub Dir1_Change()

End Sub

Private Sub optAbajo_Click()
txtCaja.Top = 4080
End Sub

Private Sub optAmarillo_Click()
txtCaja.BackColor = vbYellow
End Sub

Private Sub optArriba_Click()
txtCaja.Top = 0

End Sub

Private Sub optAzul_Click()

txtCaja.BackColor = vbBlue
End Sub

Private Sub optRojo_Click()
txtCaja.BackColor = vbRed
End Sub

Private Sub optVerde_Click()
txtCaja.BackColor = vbGreen
End Sub

