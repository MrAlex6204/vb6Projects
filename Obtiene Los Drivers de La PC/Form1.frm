VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long

Private Sub Aceptar_Click()
Dim Ret As Long
Dim i As Long
  
'Esta variable va almacenando los drives
Dim Las_Unidades As String
  
  
Ret = GetLogicalDrives
  
Las_Unidades = "Drives disponibles: "
  
For i = 0 To 25
  
    If (Ret And 2 ^ i) <> 0 Then
        Las_Unidades = Las_Unidades & " " + Chr$(65 + i)
    End If
  
Next i
  
'Mostramos el String que contiene las unidades
MsgBox Las_Unidades, vbInformation
  
End Sub
  

