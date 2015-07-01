VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Captura de Articulos"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12165
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Width           =   7935
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCTO"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PRECIO"
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
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   4080
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Captura de Articulos"
      BeginProperty Font 
         Name            =   "Asimov"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   4125
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


 
 pRECIO(i) = Text2
 NOMBREPRODUCTO(i) = Text1
 i = i + 1


End Sub



Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Form_Load()
i = 0
End Sub

