VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   3870
   ClientLeft      =   3465
   ClientTop       =   3465
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogoTipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7320
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   2880
         TabIndex        =   8
         Top             =   3120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1085
         _Version        =   327682
         Appearance      =   0
         Max             =   10000
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3240
         Top             =   2160
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmLogoTipo.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2055
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00008080&
         Caption         =   "Copyright:Oscar A. Vera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   2580
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00008080&
         Caption         =   "Company:Comunicable S.A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   2790
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00008080&
         Caption         =   "VeraSoftWare Develoment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         TabIndex        =   4
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Net Chat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5565
         TabIndex        =   5
         Top             =   1860
         Width           =   1290
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "VeraSoft"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   7
         Top             =   1140
         Width           =   2655
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "U.A.T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Sub ProgressBar1_Click()
Timer1.Enabled = False
Dim i As Integer
i = 0
While i < ProgressBar1.Max
ProgressBar1.Value = i
i = i + 1
Wend
 Unload Me
frmLogin.Show

End Sub

Private Sub Timer1_Timer()
Call ProgressBar1_Click
End Sub
