VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar grafico"
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3015
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   120
         Pattern         =   "*.bmp;*.gif;*.jpg;*.jpeg"
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   5
      Top             =   4680
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   3375
   End
   Begin VB.PictureBox Picture2 
      Height          =   3255
      Left            =   6840
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   3240
      ScaleHeight     =   3195
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pi180 = 3.14159265358979 / 180
Private Declare Function SetStretchBltMode Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nStretchMode As Long) As Long

Private Declare Function PlgBlt Lib "gdi32.dll" ( _
    ByVal hdcDest As Long, _
    lpPicture1int As Picture1INTAPI, _
    ByVal hdcSrc As Long, _
    ByVal nXSrc As Long, _
    ByVal nYSrc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hbmMask As Long, _
    ByVal xMask As Long, _
    ByVal yMask As Long) As Long

Private Type Picture1INTAPI
        x As Long
        y As Long
End Type

Dim PtList(2) As Picture1INTAPI



Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub File1_Click()
CargarImagen File1.Path & "\" & File1.FileName
End Sub

Private Sub CargarImagen(Path As String)
Dim i As Integer

For i = 0 To 2
With HScroll1(i)
.LargeChange = 500
.Max = 180
.Min = -180
End With
Next

With HScroll1(3)
.LargeChange = 500
.Max = 86
.Min = 1
.Value = 45
End With

With Picture2
.AutoSize = True
.Visible = False
.AutoRedraw = True
.ScaleMode = vbPixels
.Picture = LoadPicture(Path)

End With
With Picture1
.AutoRedraw = True
.ScaleMode = vbPixels
'.Picture = LoadPicture("C:\WINDOWS\Pompas.bmp")
.AutoSize = True
End With

Call DoRedraw
End Sub



Sub DoRedraw()
    Dim x As Integer
    Dim NewX As Integer, NewY As Integer
    Dim SinAng1, CosAng1, SinAng2, SinAng3
    Dim Zoom

    'Punktliste zurücksetzen:
    PtList(0).x = -(Picture2.ScaleWidth / 2)
    PtList(0).y = -(Picture2.ScaleHeight / 2)
    PtList(1).x = Picture2.ScaleWidth / 2
    PtList(1).y = -(Picture2.ScaleHeight / 2)
    PtList(2).x = -(Picture2.ScaleWidth / 2)
    PtList(2).y = (Picture2.ScaleHeight / 2)
    
    
    Zoom = Tan(HScroll1(3).Value * pi180)
    SinAng1 = Sin((HScroll1(0).Value + 90) * pi180)
    CosAng1 = Cos((HScroll1(0).Value + 90) * pi180)
    SinAng2 = Sin((HScroll1(1).Value + 90) * pi180) * Zoom
    SinAng3 = Sin((HScroll1(2).Value + 90) * pi180) * Zoom
    
    For x = 0 To 2
        NewX = (PtList(x).x * SinAng1 + PtList(x).y * CosAng1) * SinAng2
        NewY = (PtList(x).y * SinAng1 - PtList(x).x * CosAng1) * SinAng3
        PtList(x).x = NewX + (Picture1.ScaleWidth / 2)
        PtList(x).y = NewY + (Picture1.ScaleHeight / 2)
    Next
    
    Picture1.Cls
    'opcional-- suaviza un poco cuando esta la figura en posocion recta
    SetStretchBltMode Picture1.hdc, vbPaletteModeNone
    'api que rota la imagen
    Call PlgBlt(Picture1.hdc, PtList(0), Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, 0, 0, 0)
    Picture1.Refresh

End Sub

Private Sub HScroll1_Scroll(Index As Integer)
Call DoRedraw
End Sub
