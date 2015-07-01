VERSION 5.00
Begin VB.Form frmScroll 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marquesina. Scroll horizontal de texto"
   ClientHeight    =   1095
   ClientLeft      =   3780
   ClientTop       =   2910
   ClientWidth     =   4860
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "scrolltext.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1095
   ScaleWidth      =   4860
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdScroll 
      Caption         =   "Iniciar scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   600
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strText As String

Private Sub cmdScroll_Click()
    strText = String(30, " ") + "VeraSoft Development"
    
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    strText = Mid(strText, 2) & Left(strText, 1)
    txtMessage = strText
End Sub





