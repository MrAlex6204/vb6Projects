VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   3360
      TabIndex        =   5
      Top             =   5880
      Width           =   2895
   End
   Begin VB.PictureBox pctBox 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2280
      ScaleHeight     =   3075
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtCaja3 
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Text            =   "0"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtCaja2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.HScrollBar hsbX 
      Height          =   495
      LargeChange     =   5
      Left            =   2160
      TabIndex        =   1
      Top             =   4320
      Width           =   6135
   End
   Begin VB.VScrollBar vsbY 
      Height          =   3855
      LargeChange     =   5
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
pctBox.Scale (0, 0)-(100, 100)
End Sub

Private Sub HScroll1_Change()
txtCaja3.Text = Format(hsbX.Value)
pctBox.PSet (hsbX.Value, vsbY.Value), vbRed
End Sub

Private Sub hsbX_Change()
txtCaja3.Text = Format(hsbX.Value)
pctBox.PSet (hsbX.Value, vsbY.Value), vbBlack
End Sub

Private Sub txtCaja2_KeyPress(KeyAscii As Integer)
Dim valor As Integer
valor = Val(txtCaja2.Text)
If KeyAscii = 13 Then
If valor <= vsbY.Max And valor >= vsbY.Min Then
vsbY.Value = valor
ElseIf valor > vsbY.Max Then
vsbY.Value = vsbY.Max
Else
vsbY.Value = vsbY.Min
End If
End If
End Sub

Private Sub txtCaja3_KeyPress(KeyAscii As Integer)
Dim valor As Integer
valor = Val(txtCaja3.Text)
If KeyAscii = 13 Then
If valor <= hsbX.Max And valor >= hsbX.Min Then
hsbX.Value = valor
ElseIf valor > hsbX.Max Then
hsbX.Value = hsbX.Max
Else
hsbX.Value = hsbX.Min
End If
End If
End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub vsbY_Change()
txtCaja2.Text = Format(vsbY.Value)
pctBox.PSet (hsbX.Value, vsbY.Value), vbBlue
End Sub
