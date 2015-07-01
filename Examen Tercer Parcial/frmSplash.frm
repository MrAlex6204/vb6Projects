VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4935
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar Bar 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
      Max             =   1e6
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7080
      Top             =   960
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8235
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Timer1_Timer()

Dim i As Long
While i < 1000000
Bar.Value = i
i = i + 1
Wend

Timer1.Enabled = False
Unload Me
MDIForm1.Show
End Sub
