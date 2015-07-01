VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "Form1"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdconvercion2 
      Caption         =   ">>"
      Height          =   855
      Left            =   4800
      TabIndex        =   8
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtdecimal2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox txtbinario2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox txtbinario1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtdecimal1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2760
      Width           =   3735
   End
   Begin VB.CommandButton cmdconvercion1 
      Caption         =   ">>"
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Numero Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Numero Binario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   7320
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Numero Binario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Numero Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Convercion Binario  a Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Convercion Decimal  a Binario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Convercion de Binario a Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub convercion1_Click()


   
   
   







End Sub

Private Sub cmdconvercion1_Click()
dimnumero As Long, numdec As Long
Dim numbin As String * 20, bit As String * 1
Dim division As Boolean
Dim residuo As Byte
Dim checar As Boolean

numdec = Val(txtdecimal1.Text)
numero = numdec
checar = False

Do
   Do While numero <> 0
   division = numero / 2
   residuo = numero - (division * 2)
   bit = CStr(residuo)
   numbit = bit + numbin
   numero = division
   Loop
   checar = False
   Loop Until checar = True
   txtbinario1.Text = numbin
   
   
End Sub

Private Sub cmdconvercion2_Click()

End Sub
