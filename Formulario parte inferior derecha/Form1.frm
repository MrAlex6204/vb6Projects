VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
' Position the form in the lower right corner.
Private Sub PutFormInLowerRight(ByVal frm As Form, ByVal right_margin As Single, ByVal bottom_margin As Single)
Dim wa_info As RECT

    If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then
    
        ' We got the information. Position the form.
        ' Position the form.
        frm.Left = ScaleX(wa_info.Right, vbPixels, vbTwips) - Width - right_margin
        
        frm.Top = ScaleY(wa_info.Bottom, vbPixels, vbTwips) - Height - bottom_margin
    Else
        ' We did not get the work area bounds.
        ' Use the entire screen.
        frm.Left = Screen.Width - Width - right_margin
        frm.Top = Screen.Height - Height - bottom_margin
    End If
End Sub


Private Sub Form_Load()
    PutFormInLowerRight Me, 0, 0
End Sub


