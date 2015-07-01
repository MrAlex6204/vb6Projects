VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "www.activevb.de"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Beenden"
      Height          =   675
      Left            =   2340
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit



Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\Back.gif")
    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
    
    MakeFormTransparent Me, vbMagenta
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
