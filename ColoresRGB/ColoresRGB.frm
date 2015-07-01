VERSION 5.00
Begin VB.Form frmColores 
   Caption         =   "Colores"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   2
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   1
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   0
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "0"
      Top             =   3360
      Width           =   855
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   2
      Left            =   1200
      Max             =   255
      TabIndex        =   7
      Top             =   4560
      Width           =   5055
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   1
      Left            =   1200
      Max             =   255
      TabIndex        =   6
      Top             =   3960
      Width           =   5055
   End
   Begin VB.HScrollBar hsbColor 
      Height          =   375
      Index           =   0
      Left            =   1200
      Max             =   255
      TabIndex        =   5
      Top             =   3360
      Width           =   5055
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton optColor 
      Caption         =   "Fondo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label lblColores 
      Caption         =   "Azul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   13
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblColores 
      Caption         =   "Verde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblColores 
      Caption         =   "Rojo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblCuadro 
      Caption         =   "INFORMATICA 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "frmColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Brojo, Bverde, Bazul As Integer
Public Frojo, Fverde, Fazul As Integer
Private Sub cmdSalir_Click()
End
End Sub
Private Sub Form_Load()

Brojo = 0
Bverde = 0
Bazul = 0
Frojo = 255
Fverde = 255
Fazul = 255
lblCuadro.BackColor = RGB(Brojo, Bverde, Bazul)
lblCuadro.ForeColor = RGB(Frojo, Fverde, Fazul)

End Sub
Private Sub hsbColor_Change(Index As Integer)
If optColor(0).Value = True Then
lblCuadro.BackColor = RGB(hsbColor(0).Value, hsbColor(1).Value, _
hsbColor(2).Value)
Dim i As Integer
For i = 0 To 2
txtColor(i).Text = hsbColor(i).Value
Next i
Else
lblCuadro.ForeColor = RGB(hsbColor(0).Value, hsbColor(1).Value, _
hsbColor(2).Value)
For i = 0 To 2
txtColor(i).Text = hsbColor(i).Value
Next i
End If
End Sub

Private Sub optColor_Click(Index As Integer)
If Index = 0 Then 'Se pasa a cambiar el fondo
Frojo = hsbColor(0).Value
Fverde = hsbColor(1).Value
Fazul = hsbColor(2).Value
hsbColor(0).Value = Brojo
hsbColor(1).Value = Bverde
hsbColor(2).Value = Bazul
Else 'Se pasa a cambiar el texto
Brojo = hsbColor(0).Value
Bverde = hsbColor(1).Value
Bazul = hsbColor(2).Value
hsbColor(0).Value = Frojo
hsbColor(1).Value = Fverde
hsbColor(2).Value = Fazul
End If

End Sub

