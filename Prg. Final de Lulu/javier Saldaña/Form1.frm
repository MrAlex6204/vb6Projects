VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Captura de Articulos"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Precio:"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   555
      Left            =   5880
      TabIndex        =   0
      Top             =   1200
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Command1_Click()

If i > 10 Then
MsgBox "No hay Espacio Para mas Articulos", vbCritical
Exit Sub
Else
 
 Precio(i) = Text2
 NombreProduc(i) = Text1
 i = i + 1
End If
Text1 = Empty
Text2 = Empty
End Sub

Private Sub Command2_Click()
Text1.Text = Empty
Text2.Text Empty
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
End Sub

