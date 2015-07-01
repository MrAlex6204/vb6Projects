VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2190
   ClientLeft      =   9195
   ClientTop       =   8715
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Label1.Caption = "Width: " + Str(Screen.Width) + " Heigth: " + Str(Screen.Height)
Label2.Caption = "Width:" + Str(Dialog.Width) + " Heigth:" + Str(Dialog.Height) + " Left:" + Str(Dialog.Left) + " Top:" + Str(Dialog.Top)
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Label2.Caption = "Width:" + Str(Dialog.Width) + " Heigth:" + Str(Dialog.Height) + " Left:" + Str(Dialog.Left) + " Top:" + Str(Dialog.Top)
End Sub

Private Sub OKButton_Click()
'FRM.Move Screen.Width - FRM.Width, Screen.Height - FRM.Height - Tamaño - FormAbierto
Dialog.Move Screen.Width - Dialog.Width, Screen.Height - Dialog.Height
End Sub
