VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Clone"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Trim (Text1)
If Text1.Text = "" Then
MsgBox "Please Enter a Number."
GoTo 1
Else: GoTo 2
End If
2:
Dim a As Integer
a = Text1.Text
For i = 1 To (a)
Form1.Caption = (i)
Set Form1 = New Form1
Form1.Visible = True
Next i
1:
End Sub

Private Sub Form_Load()
Command1.Caption = "Type A Number In The Text Box"
End Sub

Private Sub Text1_Change()
Command1.Caption = "&Open " & Text1.Text + " New Form(s)"
If IsNumeric(Text1.Text) = False Then
   Text1.Text = ""
   End If
End Sub
