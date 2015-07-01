VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5175
   ClientLeft      =   4920
   ClientTop       =   3570
   ClientWidth     =   4125
   ControlBox      =   0   'False
   DrawMode        =   10  'Mask Pen
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   206.25
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mover Form"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Form Top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   1020
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   360
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   360
      Picture         =   "Form1.frx":2798
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Top "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1290
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   360
      Picture         =   "Form1.frx":4F45
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   -360
      Picture         =   "Form1.frx":74A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Dim i, j As Long
i = Form2.Top
j = Form2.Left
While ((i < Screen.Height) = (j < Screen.Width - Form1.Width)) = True
Form2.Top = i
Form2.Left = j
i = i + 1
j = j + 1
Wend
Form1.Top = 0

End Sub


 Sub Command3_Click()
Image1.Picture = Image2.Picture

Text1.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text2.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text3.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text4.BackColor = &H40C0&: Text1.ForeColor = vbBlack

Label1.ForeColor = vbBlack
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

 Sub Command4_Click()
Image1.Picture = Image4.Picture

Text1.BackColor = vbBlack
Text2.BackColor = vbBlack
Text3.BackColor = vbBlack
Text4.BackColor = vbBlack

Label1.ForeColor = &HC0C0C0
Label2.ForeColor = &HC0C0C0
Label3.ForeColor = &HC0C0C0
Label4.ForeColor = &HC0C0C0
End Sub

 Sub Command5_Click()
Image1.Picture = Image3.Picture

Text1.BackColor = vbBlack
Text2.BackColor = vbBlack
Text3.BackColor = vbBlack
Text4.BackColor = vbBlack

Label1.ForeColor = vbBlack
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

      Private Sub Form_Load()
         Command1.Caption = "Exit"
         
         Form2.Top = Form1.Top + 30
         Form2.Left = Form1.Left + Form1.Width
          mover.MoverForm
         
      End Sub



      Private Sub Command1_Click()
        Unload Me
      End Sub





Private Sub Timer1_Timer()
mover.MoverForm
End Sub
