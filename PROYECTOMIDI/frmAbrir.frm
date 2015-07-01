VERSION 5.00
Begin VB.Form frmAbrir 
   Caption         =   "Abrir"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      Height          =   2400
      ItemData        =   "frmAbrir.frx":0000
      Left            =   960
      List            =   "frmAbrir.frx":0002
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "&Abrir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nom. Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmAbrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Integer

Private Sub cmdAbrir_Click()
Open "Temp.txt" For Input As #1

Dim Filenom As String
Dim value As Integer

Do While Not EOF(1)
Line Input #1, Filenom


lstFiles.AddItem (Filenom)

i = i + 1
Loop
Close #1

End Sub

