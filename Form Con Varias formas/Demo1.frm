VERSION 5.00
Object = "{40BD39E3-6F1E-11D1-B2DF-444553540000}#1.0#0"; "SHAPE.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3732
   LinkTopic       =   "Form1"
   Picture         =   "Demo1.frx":0000
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   270
      Left            =   105
      TabIndex        =   1
      Top             =   3675
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   165
      Left            =   3345
      Picture         =   "Demo1.frx":10BC2
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   405
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.jcomsoft.com   jinhui@jcomsoft.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   585
      TabIndex        =   5
      Top             =   2370
      Width           =   2670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Made by Jin Hui"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   570
      TabIndex        =   4
      Top             =   1770
      Width           =   2670
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Shape ActiveX 1.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   570
      TabIndex        =   3
      Top             =   855
      Width           =   2670
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shape"
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   225
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin FormShape.FormShape FormShape1 
      Left            =   2880
      Top             =   3480
      ShapeType       =   1
      MaskColor       =   65280
      AutoScale       =   0   'False
      ScaleX          =   1
      ScaleY          =   1
      ShapeString     =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oldleft As Integer
Public oldtop As Integer
Public oldx As Integer
Public oldy As Integer
Public moving As Boolean

Private Sub Command1_Click()
    End
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.SetFocus
End Sub

Private Sub Form_Load()
    rx = Form1.Width / Screen.TwipsPerPixelX / 248
    ry = Form1.Height / Screen.TwipsPerPixelY / 272
    Form1.Width = Form1.Width / rx
    Form1.Height = Form1.Height / ry
    Command1.Left = Command1.Left / rx
    Command1.Top = Command1.Top / rx
    Command1.Width = Command1.Width / rx
    Command1.Height = Command1.Height / ry
    Label1.Left = Label1.Left / rx
    Label1.Top = Label1.Top / ry
    Label2.Left = Label2.Left / rx
    Label2.Top = Label2.Top / ry
    Label2.Width = Label2.Width / rx
    Label2.Font.Size = Label2.Font.Size / rx
    Label3.Left = Label3.Left / rx
    Label3.Top = Label3.Top / ry
    Label3.Width = Label3.Width / rx
    Label3.Font.Size = Label3.Font.Size / rx
    Label4.Left = Label4.Left / rx
    Label4.Top = Label4.Top / ry
    Label4.Width = Label4.Width / rx
    Label4.Font.Size = Label4.Font.Size / rx
    FormShape1.hWnd = Form1.hWnd
    FormShape1.ShapePicture = Form1.Picture
    moving = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 40 And Y > 3 Then
        moving = True
        oldleft = Form1.Left
        oldtop = Form1.Top
        oldx = X * Screen.TwipsPerPixelX + Form1.Left
        oldy = Y * Screen.TwipsPerPixelY + Form1.Top
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If moving Then
        thisx = X * Screen.TwipsPerPixelX + Form1.Left
        thisy = Y * Screen.TwipsPerPixelY + Form1.Top
        Form1.Left = oldleft + thisx - oldx
        Form1.Top = oldtop + thisy - oldy
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = False
End Sub

