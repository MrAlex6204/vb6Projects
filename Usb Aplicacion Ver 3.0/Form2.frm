VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5745
   ClientLeft      =   4920
   ClientTop       =   3570
   ClientWidth     =   7770
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   287.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   388.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Ayuda?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   6480
      MaskColor       =   &H000000C0&
      Picture         =   "Form2.frx":08E2
      TabIndex        =   5
      Top             =   80
      Width           =   975
   End
   Begin VB.ComboBox Color_Texto 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   420
      ItemData        =   "Form2.frx":2094
      Left            =   1800
      List            =   "Form2.frx":20A1
      TabIndex        =   4
      Text            =   "Skin"
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Examinar"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4920
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "Solo Inmagenes"
      DialogTitle     =   "Usb Aplication VeraSoft Develoment"
      Filter          =   "Imágenes(*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
   End
   Begin VB.Image Pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   600
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   6000
      Picture         =   "Form2.frx":20CA
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   6000
      Picture         =   "Form2.frx":4627
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   6120
      Picture         =   "Form2.frx":6DBF
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5775
      Left            =   0
      Picture         =   "Form2.frx":956C
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








Private Sub Color_Texto_Change()

If Color_Texto.Text = "Skin Naranja" Then
Image1.Picture = Image2.Picture
Call Form2.Command3_Click
End If

If Color_Texto.Text = "Skin Gris" Then
Image1.Picture = Image3.Picture
Call Form2.Command5_Click
End If

If Color_Texto.Text = "Skin Negro" Then
Image1.Picture = Image4.Picture
Call Form2.Command4_Click
End If

End Sub

Private Sub Color_Texto_Click()
If Color_Texto.Text = "Skin Naranja" Then
Image1.Picture = Image2.Picture
Call Form2.Command3_Click
End If

If Color_Texto.Text = "Skin Gris" Then
Image1.Picture = Image3.Picture
Call Form2.Command5_Click
End If

If Color_Texto.Text = "Skin Negro" Then
Image1.Picture = Image4.Picture
Call Form2.Command4_Click
End If
End Sub

Private Sub Color_Texto_KeyDown(KeyCode As Integer, Shift As Integer)
If Color_Texto.Text = "Skin Naranja" Then
Image1.Picture = Image2.Picture
Call Form2.Command3_Click
End If

If Color_Texto.Text = "Skin Gris" Then
Image1.Picture = Image3.Picture
Call Form2.Command5_Click
End If

If Color_Texto.Text = "Skin Negro" Then
Image1.Picture = Image4.Picture
Call Form2.Command4_Click
End If
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Text1.Text = CommonDialog1.FileName
'Cargamos la imagen del path que tiene text1

End If
'Image1.Picture = LoadPicture(Text1.Tex)
End Sub

Private Sub Command3_Click()
frmAbout.Show
End Sub

Private Sub Command6_Click()
Form2.Show


 ver = False
End Sub
Private Sub Form_Load()
Dim i As Integer
         
        
        
        
End Sub
     Private Sub Command1_Click()
         End
      End Sub
Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngReturnValue As Long


        If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Form1.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
        Mover.MoverForm
        End If
        
        
       
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long


        If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Form1.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
        Mover.MoverForm
        End If
        
End Sub

Private Sub Text1_Change()
Pic.Picture = LoadPicture(Text1)
End Sub

 Sub Timer1_Timer()
   Form2.Top = Form1.Top + 30
   Form2.Left = Form1.Left + Form1.Width
End Sub
