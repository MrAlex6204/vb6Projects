VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Splitter"
   ClientHeight    =   5145
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MousePointer    =   7  'Size N S
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5145
   ScaleWidth      =   4965
   Begin VB.PictureBox picSE 
      BackColor       =   &H0080C0FF&
      Height          =   1815
      Left            =   2160
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
      Begin VB.Label lblSE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SouthEast"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.PictureBox picNE 
      BackColor       =   &H00FFFF80&
      Height          =   1815
      Left            =   2160
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      Begin VB.Label lblNE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NorthEast"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.PictureBox picSW 
      BackColor       =   &H008080FF&
      Height          =   1815
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
      Begin VB.Label lblSW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SouthWest"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2040
      End
   End
   Begin VB.PictureBox picNW 
      BackColor       =   &H0080FF80&
      Height          =   1815
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   600
      Width           =   1815
      Begin VB.Label lblNW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NorthWest"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SPLITTER_HEIGHT = 40
Private Const SPLITTER_WIDTH = 40

' Minimum allowed sizes for any PictureBox.
Private Const MIN_HORIZONTAL = 10
Private Const MIN_VERTICAL = 10

' Drag directions.
Private Const DRAG_NONE = 0
Private Const DRAG_HORIZONTAL = 1
Private Const DRAG_VERTICAL = 2
Private Const DRAG_BOTH = 3

' The percentage occupied by the PictureBox.
Private PercentageHorizontal As Single
Private PercentageVertical As Single

' True when we are dragging the splitter.
Private DragDir As Integer
' Arrange the controls on the form.
Private Sub ArrangeControls()
Dim hgt1 As Single
Dim hgt2 As Single
Dim wid1 As Single
Dim wid2 As Single

    ' Don't bother if we're iconized.
    If WindowState = vbMinimized Then Exit Sub

    hgt1 = (ScaleHeight - SPLITTER_HEIGHT) * PercentageHorizontal
    wid1 = (ScaleWidth - SPLITTER_WIDTH) * PercentageVertical
    hgt2 = (ScaleHeight - SPLITTER_HEIGHT) - hgt1
    wid2 = (ScaleWidth - SPLITTER_WIDTH) - wid1
    picNW.Move 0, 0, wid1, hgt1
    picSW.Move 0, hgt1 + SPLITTER_HEIGHT, wid1, hgt2
    picNE.Move wid1 + SPLITTER_WIDTH, 0, wid2, hgt1
    picSE.Move wid1 + SPLITTER_WIDTH, hgt1 + SPLITTER_HEIGHT, wid2, hgt2
End Sub

' Return drag flags indicating the area at this point.
Private Function OverDragArea(ByVal X As Single, ByVal Y As Single)
Dim over As Integer

    over = DRAG_NONE
    If X > picNW.Width And X < picNE.Left _
        Then over = over Or DRAG_HORIZONTAL
    If Y > picNW.Height And Y < picSW.Top _
        Then over = over Or DRAG_VERTICAL
    OverDragArea = over
End Function

Private Sub Form_Load()
Dim i As Integer
Dim txt As String

    ' Start with the split in the middle.
    PercentageHorizontal = 0.5
    PercentageVertical = 0.75
    
    txt = ""
    For i = 1 To 5
        txt = txt & Space$(i * 4) & "NorthWest" & vbCrLf
    Next i
    lblNW.AutoSize = True
    lblNW.Caption = txt

    txt = ""
    For i = 1 To 5
        txt = txt & Space$(i * 4) & "NorthEast" & vbCrLf
    Next i
    lblNE.AutoSize = True
    lblNE.Caption = txt

    txt = ""
    For i = 1 To 5
        txt = txt & Space$(i * 4) & "SouthWest" & vbCrLf
    Next i
    lblSW.AutoSize = True
    lblSW.Caption = txt

    txt = ""
    For i = 1 To 5
        txt = txt & Space$(i * 4) & "SouthEast" & vbCrLf
    Next i
    lblSE.AutoSize = True
    lblSE.Caption = txt
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragDir = OverDragArea(X, Y)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim over As Integer

    If DragDir = DRAG_NONE Then
        ' Set the correct arrow type.
        over = OverDragArea(X, Y)
        Select Case over
            Case DRAG_HORIZONTAL
                MousePointer = vbSizeWE
            Case DRAG_VERTICAL
                MousePointer = vbSizeNS
            Case DRAG_BOTH
                MousePointer = vbSizeNWSE
        End Select
        Exit Sub
    End If

    If DragDir And DRAG_VERTICAL Then
        If Y < MIN_VERTICAL Then
            Y = MIN_VERTICAL
        ElseIf ScaleHeight - Y < MIN_VERTICAL Then
            Y = ScaleHeight - MIN_VERTICAL
        End If
        PercentageHorizontal = Y / ScaleHeight
    End If

    If DragDir And DRAG_HORIZONTAL Then
        If X < MIN_HORIZONTAL Then
            X = MIN_HORIZONTAL
        ElseIf ScaleWidth - X < MIN_HORIZONTAL Then
            X = ScaleWidth - MIN_HORIZONTAL
        End If
        PercentageVertical = X / ScaleWidth
    End If

    ArrangeControls
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragDir = DRAG_NONE
End Sub


Private Sub Form_Resize()
    ArrangeControls
End Sub
