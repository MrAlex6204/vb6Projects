VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Cordenadas del Form"
      Height          =   1695
      Left            =   6360
      TabIndex        =   22
      Top             =   360
      Width           =   5895
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   24
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cord. Y:"
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
         Left            =   1200
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cord. X:"
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
         Left            =   1200
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cursor Cordenadas"
      Height          =   2775
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   5895
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   21
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Mouse Clic:"
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
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cord. X:"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cord. Y:"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Teclado Extendido"
      Height          =   2175
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   6015
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   27
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Shift + Code:"
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
         TabIndex        =   28
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Telca Shitf:"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tecla Code:"
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
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teclado Estandar"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   1
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telca Num:"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Telca Caracter:"
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
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Telca Hexadecimal:"
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
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telca Octal:"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   1245
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
End
End Sub

Private Sub Form_Click()

Load Form2
Form2.Top = Text5
Form2.Left = Text6
Form2.Show

End Sub





'Eventos del Teclado en Forma Estandar


Private Sub Form_KeyPress(KeyAscii As Integer)
Text1.Text = Str(KeyAscii)
Text2.Text = Chr(KeyAscii)
Text3 = Hex(KeyAscii)
Text4 = Oct(KeyAscii)

End Sub
'Eventos del Teclado en forma Extendida
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Text7 = KeyCode
Text8 = Shift
Text3 = KeyCode + Shift
End Sub



'Eventos del Mouse
Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text6 = X
Text5 = Y
Text9 = Str(Button)
If Button = 0 Then
Label10.Caption = "Clic Suelto"
End If

If Button = 1 Then
Label10.Caption = "Arrastrando con Clic Izquierdo"
End If
If Button = 2 Then
Label10.Caption = "Arrastrando con Clic Derecho"
End If
If Button = 4 Then
Label10.Caption = "Arrastrando con Ruedita"
End If



End Sub


