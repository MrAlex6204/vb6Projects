VERSION 5.00
Begin VB.Form frmMinical 
   BackColor       =   &H80000007&
   Caption         =   "Calculadora Turbo"
   ClientHeight    =   3510
   ClientLeft      =   4875
   ClientTop       =   3135
   ClientWidth     =   6795
   Icon            =   "frmMinical.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdDivi 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdResta 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSuma 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtOper2 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtOper1 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Controles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label lbIgual 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lbOper 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmMinical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public Numero1 As Integer
Public Numero2 As Integer
Public Resultado As Integer



Private Sub cmdDivi_Click()
lbOper.Caption = "/"


Numero1 = Int(txtOper1.Text)
Numero2 = Int(txtOper2.Text)
txtResult.Text = Int(Numero1) / Int(Numero2)

End Sub

Private Sub cmdMulti_Click()
lbOper.Caption = "*"
Numero1 = Int(txtOper1.Text)
Numero2 = Int(txtOper2.Text)
txtResult.Text = Int(Numero1) * Int(Numero2)
End Sub

Private Sub cmdResta_Click()
lbOper.Caption = "-"
Numero1 = Int(txtOper1.Text)
Numero2 = Int(txtOper2.Text)
txtResult.Text = Int(Numero1) - Int(Numero2)
End Sub

Private Sub cmdSuma_Click()
lbOper.Caption = "+"
Numero1 = Int(txtOper1.Text)
Numero2 = Int(txtOper2.Text)
txtResult.Text = Int(Numero1) + Int(Numero2)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub
