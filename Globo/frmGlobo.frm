VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Globo"
   ClientHeight    =   5535
   ClientLeft      =   2940
   ClientTop       =   2910
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Totales"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "frmGlobo.frx":0000
      Left            =   360
      List            =   "frmGlobo.frx":000A
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mensaje"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Posicion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()


Globo.Mensaje (Text1.Text)
End Sub

Private Sub Command2_Click()

Dim i As Integer

Label1.Caption = "Totales:" + Str(List1.ListCount)
Label2.Caption = "Posicion:"

End Sub

Private Sub Form_Load()
List1.AddItem "Lista"
List1.AddItem "Oscar"
List1.AddItem "Selene"
List1.AddItem "Everad"
List1.AddItem "Alvino"
List1.AddItem "Santiago"
If (Len(Text1.Text)) = 0 Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

End Sub

Private Sub List1_Click()
Dim Pos As Integer


End Sub

Private Sub Text1_Change()
If (Len(Text1.Text)) = 0 Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
