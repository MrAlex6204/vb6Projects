VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2676
   ClientLeft      =   1140
   ClientTop       =   1512
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2676
   ScaleWidth      =   6360
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   4320
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2364
      ScaleWidth      =   6084
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Xpos As New Collection
Private Ypos As New Collection
Private BlinkText As New Collection

' Print a line, saving the position of the text
' target for blinking.
Private Sub PrintLine(txt As String, target As String)
Dim pos1 As Integer
Dim pos2 As Integer
Dim outtxt As String
Dim txtlen As Integer

    txtlen = Len(txt)
    pos1 = 1
    Do
        ' Find the next target.
        pos2 = InStr(pos1, txt, target)
        If pos2 = 0 Then
            ' Print the rest of the string and stop.
            outtxt = Mid$(txt, pos1, txtlen - pos1 + 1)
            Picture1.Print outtxt
            Exit Do
        End If
        
        ' Print the text before the target.
        ' Note semi-colon.
        outtxt = Mid$(txt, pos1, pos2 - pos1)
        Picture1.Print outtxt;

        ' If that's the end of the string, stop.
        If pos2 = 0 Then Exit Do
        
        ' Save the position.
        Xpos.Add Picture1.CurrentX
        Ypos.Add Picture1.CurrentY
        BlinkText.Add target
        
        ' Print the target.
        Picture1.Print target;

        ' Look for the next one.
        pos1 = pos2 + Len(target)
    Loop
End Sub

Private Sub Form_Load()
    Picture1.Font.bold = True
    PrintLine "Name                Date of Birth           Status", "XXXXX"
    Picture1.Font.bold = False
    PrintLine "Jane Smith          01/01/1961              Active", "Active"
    PrintLine "Rod Stephens        02/02/1962              Inactive", "Active"
    PrintLine "Julia Keeley        03/03/1963              Active", "Active"
    PrintLine "Micky Johnson       04/04/1964              Inactive", "Active"
    PrintLine "Ben Pierce          05/05/1965              Inactive", "Active"
    PrintLine "Linda Roberts       06/06/1966              Active", "Active"
End Sub

Private Sub tmrBlink_Timer()
Dim i As Integer

    If Picture1.ForeColor = vbRed Then
        Picture1.ForeColor = vbBlack
    Else
        Picture1.ForeColor = vbRed
    End If
    
    For i = 1 To Xpos.Count
        Picture1.CurrentX = Xpos(i)
        Picture1.CurrentY = Ypos(i)
        Picture1.Print BlinkText(i)
    Next i
End Sub


