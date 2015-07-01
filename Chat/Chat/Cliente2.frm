VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   3435
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Text            =   "Escrive tu Nick"
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2790
      TabIndex        =   0
      Top             =   735
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Poner a la escucha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "Cliente2.frx":0000
      Top             =   240
      Width           =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NICK:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Menu mnuMascaras 
      Caption         =   "PopMnu"
      Visible         =   0   'False
      Begin VB.Menu Mascara 
         Caption         =   "Clasico"
         Index           =   0
      End
      Begin VB.Menu Mascara 
         Caption         =   "Madera"
         Index           =   1
      End
      Begin VB.Menu Mascara 
         Caption         =   "Metal"
         Index           =   2
      End
      Begin VB.Menu Mascara 
         Caption         =   "Marmol"
         Index           =   3
      End
      Begin VB.Menu Mascara 
         Caption         =   "Ladrillo"
         Index           =   4
      End
   End
   Begin VB.Menu MnuMostrar 
      Caption         =   "Mostrar"
      Visible         =   0   'False
      Begin VB.Menu PopMnuMostar 
         Caption         =   "Mostrar"
         Index           =   0
      End
      Begin VB.Menu PopMnuMostar 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu PopMnuMostar 
         Caption         =   "Salir"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Form1.Winsock1.Close
Form1.Winsock1.CONNECT Text1(0), 999
Form1.Caption = "Cliente"

End Sub

Private Sub Command2_Click()
Form1.Winsock1.Close
Form1.Winsock1.LocalPort = 999
Form1.Winsock1.Listen
Form1.Caption = "Servidor"
AgregarIcono
MostrarGlobo ("Servidor en escucha")
Me.WindowState = vbMinimized
Me.Hide

End Sub

Private Sub Command3_Click()
QuitarIcono
End
End Sub

Private Sub Form_Load()

Me.AutoRedraw = True
Dim i As Integer, Y As Integer
 For i = 0 To 350
        Me.Line (0, Y)-(Me.Width, Y + 2), RGB(0, 0, i), BF
        
        Y = Y + 10
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Form1.Visible = False Then End
QuitarIcono
End Sub

Private Sub Mascara_Click(Index As Integer)
Form1.cargarImagenes (Index)
Unload Me
End Sub
Private Sub PopMnuMostar_Click(Index As Integer)
Select Case Index
Case 0
Me.Visible = True
Me.WindowState = vbNormal

Case 2

Command3_Click
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub
