VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Captura de Vendedores"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   7935
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   3480
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4680
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Id Vendedor:"
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
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1380
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
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   555
      Left            =   6240
      TabIndex        =   6
      Top             =   1560
      Width           =   2085
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Command1_Click()

If i > 10 Then
MsgBox "No hay Espacio Para mas Vendedores", vbCritical
Exit Sub
Else

 Nombre(i) = Text1
 
 i = i + 1
End If
Text1 = Empty
Text2 = Empty
End Sub

Private Sub Command2_Click()
Text1 = Empty
Text2 = Empty
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
End Sub
