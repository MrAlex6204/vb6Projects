VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   4920
   ClientTop       =   3570
   ClientWidth     =   7845
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   392.25
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   3960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Form 2"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fondo 2"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fondo 3"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fondo 1"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mover Form"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
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
      Top             =   2040
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
      Top             =   1560
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
      Top             =   1080
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
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2640
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
      Top             =   600
      Width           =   1020
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   3720
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   3720
      Picture         =   "Form2.frx":255D
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   3840
      Picture         =   "Form2.frx":4CF5
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
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
      Top             =   2040
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
      Top             =   1560
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
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Picture         =   "Form2.frx":74A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long


      Private Declare Sub ReleaseCapture Lib "User32" ()
     
      

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2
      

Public ver  As Boolean





Private Sub Command2_Click()
Dim i, j As Long
i = Form1.Top
j = Form1.Left
While ((i < Screen.Height) = (j < Screen.Width - Form1.Width)) = True
Form1.Top = i
Form1.Left = j
i = i + 1
j = j + 1
Wend
Form1.Top = 0

End Sub


Private Sub Command3_Click()
Image1.Picture = Image2.Picture
Call Form2.Command3_Click
Text1.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text2.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text3.BackColor = &H40C0&: Text1.ForeColor = vbBlack
Text4.BackColor = &H40C0&: Text1.ForeColor = vbBlack

Label1.ForeColor = vbBlack
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Command4_Click()
Image1.Picture = Image4.Picture
Call Form2.Command4_Click
Text1.BackColor = vbBlack
Text2.BackColor = vbBlack
Text3.BackColor = vbBlack
Text4.BackColor = vbBlack

Label1.ForeColor = &HC0C0C0
Label2.ForeColor = &HC0C0C0
Label3.ForeColor = &HC0C0C0
Label4.ForeColor = &HC0C0C0
End Sub

Private Sub Command5_Click()
Image1.Picture = Image3.Picture
Call Form2.Command5_Click

Text1.BackColor = vbBlack
Text2.BackColor = vbBlack
Text3.BackColor = vbBlack
Text4.BackColor = vbBlack

Label1.ForeColor = vbBlack
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Command6_Click()
Form2.Show


 ver = False
End Sub

      Private Sub Form_Load()
Load Form2
         Command1.Caption = "Exit"
         
         Form1.Top = Screen.Height - Form1.Height
         Form1.Left = Screen.Width - Form1.Width
      End Sub
     Private Sub Command1_Click()
         End
      End Sub
Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 Dim lngReturnValue As Long
Dim lngReturnValue2 As Long

        If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Form1.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
        mover.MoverForm
        End If
             
             
       If ver = False Then
     '  Form2.Top = Form1.Top + 30
      ' Form2.Left = Form1.Left + Form1.Width
       End If
        
        
         Text1 = Form1.Top
         Text2 = Form1.Left
         Text3 = X
         Text4 = Y
End Sub

 Sub Timer1_Timer()
   Form2.Top = Form1.Top + 30
   Form2.Left = Form1.Left + Form1.Width
End Sub
