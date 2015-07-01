VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4215
   ClientLeft      =   4920
   ClientTop       =   3570
   ClientWidth     =   7710
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog RepDialogo 
      Left            =   5640
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   7440
      TabIndex        =   1
      Top             =   0
      Width           =   120
   End
   Begin WMPLibCtl.WindowsMediaPlayer Reproductor 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "VeraSoft"
      Top             =   0
      Width           =   7620
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   -1  'True
      _cx             =   13441
      _cy             =   7408
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label24_Click()

End Sub

Private Sub Form_Load()


'Me.Width = Form1.PIC.Width
'Me.Height = Form1.PIC.Height
End Sub

Private Sub lblNombre_Click()
Form1.RepVisible = False
Me.Hide
End Sub
Sub PLAYFILE()
Reproductor.Controls.Play

End Sub
Sub STOPFILE()
Reproductor.Controls.stop

End Sub
Sub PAUSEFILE()
Reproductor.Controls.Pause

End Sub

Sub Abrir()
RepDialogo.ShowOpen

If RepDialogo.FileName = "  " Then
Exit Sub
Else
Reproductor.URL = RepDialogo.FileName
End If
PLAYFILE
End Sub
Sub VOLUMEN(N As Integer)
Reproductor.settings.volume = N
End Sub


Private Sub Slider1_Change()
VOLUMEN (Slider1.Value)
End Sub

Private Sub Reproductor_DomainChange(ByVal strDomain As String)
Form1.lblNombre = strDomain
End Sub






