VERSION 5.00
Begin VB.Form frmFactorial 
   Caption         =   "Factorial"
   ClientHeight    =   4950
   ClientLeft      =   3225
   ClientTop       =   3255
   ClientWidth     =   9030
   Icon            =   "frmFactorial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   9030
   Begin VB.Timer Timer1 
      Left            =   8040
      Top             =   3240
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calculo Factorial"
      Height          =   2655
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
      Begin VB.TextBox txtResultado 
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtNumero 
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblEtiqueta 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Factorial"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Numero"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estructuras de Control"
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
      Begin VB.OptionButton optWhile 
         Caption         =   "While....Wend"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optDo 
         Caption         =   "Do...loop"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optFor 
         Caption         =   "For...Next"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Calculo Factorial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmFactorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Numero As Integer
Dim Fac As Long ' se declara lon xq con Integer no alcansa la numeracion


Private Sub optDo_Click()
Numero = Fac
Do

Loop

End Sub

Private Sub optFor_Click()
' Uso del Ciclo For...Net

Fac = Val(txtNumero.Text)
For Numero = Fac - 1 To 1 Step -1
Fac = Fac * Numero
Next Numero
txtResultado.Text = Str(Fac)
lblEtiqueta.Caption = "..Ejecucion con el For.....Next"

End Sub

