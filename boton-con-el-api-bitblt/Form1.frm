VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicImagen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   3480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   6900
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   6900
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1920
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Botón 2"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Botón 1"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Dim MouseForm As Boolean
Dim MousePicture As Boolean


Private Sub Form_Load()
Form_MouseMove 0, 0, 0, 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Normal

End Sub

Private Sub Normal()
If MouseForm = False Then
    BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 10, 15, vbSrcCopy
    BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 100, 15, vbSrcCopy
    MouseForm = True
    MousePicture = False
    refrescar
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
    refrescar

End Sub

Private Sub Picture1_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, Y As Single)

    
If MousePicture = False Then
    BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 283, 15, vbSrcCopy
    
    MousePicture = True
    MouseForm = False
    refrescar

End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Normal
MousePicture = False
End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 190, 15, vbSrcCopy
refrescar

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MousePicture = False Then
    
    BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 370, 15, vbSrcCopy
    MousePicture = True
    MouseForm = False
    refrescar

End If

End Sub

Private Sub refrescar()
    Picture1.Refresh
    Picture2.Refresh
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Normal
MousePicture = False
End Sub
