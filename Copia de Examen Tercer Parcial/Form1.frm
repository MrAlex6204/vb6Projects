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
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   7935
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descricion:"
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
         Top             =   480
         Width           =   1170
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
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Captura de Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   4380
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
 
 PrecioArticulo(i) = Text2
 NomArticulo(i) = Text1
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

