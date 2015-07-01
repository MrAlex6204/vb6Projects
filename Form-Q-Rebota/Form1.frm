VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Abajo"
      Height          =   735
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Arriba"
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Left            =   5400
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valores"
      Height          =   2775
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pantalla de Ancho:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fuerza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fuerza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pantalla de Alto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fuerza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Velocidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1110
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fuerza As Double
Dim Velocidad As Double

Dim frmFuerza As Double
Dim frmVelocidad As Double


Private Sub Command2_Click()
frmFuerza = -1
Text2 = frmFuerza
End Sub

Private Sub Command3_Click()
frmFuerza = 1
Text2 = frmFuerza
End Sub

Private Sub Form_Click()


Timer3.Enabled = False
End Sub

Private Sub Form_Load()
Label5.Caption = Screen.Height
Label6.Caption = Screen.Width

'Maximiza el formulario

Timer3.Enabled = True
Timer3.Interval = 10
'Se establece valores a las variables
frmFuerza = 0.5
frmVelocidad = 0
Text2 = frmFuerza


End Sub




Private Sub Timer3_Timer()
If Me.Top > Screen.Height - Me.Height Then
' Se ejecuta cundo llega al limite
Me.Top = Screen.Height - Me.Height
frmVelocidad = 2 * frmVelocidad * -frmFuerza
End If
frmVelocidad = frmVelocidad + frmFuerza
'mueve el top en el control
Me.Top = Me.Top + frmVelocidad
Text1 = frmVelocidad




End Sub



Private Sub VScroll2_Change()


End Sub
