VERSION 5.00
Begin VB.Form Form4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   LinkTopic       =   "Form4"
   ScaleHeight     =   2505
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   3000
      Top             =   70
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Lo estan invitando a escuchar su vos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Lo estan invitando a escuchar su vos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2940
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.EscucharSuAudio
SlideForm Me, 1
End Sub

Private Sub Command2_Click()
SlideForm Me, 1
End Sub

Private Sub Form_Load()
Image1(1).Picture = Form1.Bootons(2).Picture
End Sub

Private Sub Image1_Click(Index As Integer)
If Index = 0 Then SlideForm Me, 1
End Sub
