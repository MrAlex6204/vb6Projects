VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrowser1 
   BackColor       =   &H00000000&
   Caption         =   "VeraSoft  [Web visor]"
   ClientHeight    =   7830
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   11835
   Icon            =   "frmBrowser1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   15240
      ExtentX         =   26882
      ExtentY         =   15478
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   12120
      Top             =   600
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   6360
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":0BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":0EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":1186
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":1468
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser1.frx":1A2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   15240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   15240
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   5880
         TabIndex        =   7
         Text            =   "¯¯END!"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1530
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5775
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00C0C0C0&
         Height          =   1440
         Left            =   5880
         TabIndex        =   5
         Top             =   0
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   435
         Left            =   11640
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         Height          =   435
         Left            =   12480
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Abrir"
         Height          =   435
         Left            =   13320
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11640
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "VeraSoft"
         Filter          =   "*.*"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   555
         Left            =   0
         TabIndex        =   8
         Top             =   1560
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmBrowser1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Command4_Click()
CommonDialog1.ShowOpen
End Sub

Private Sub Command1_Click()
On Error Resume Next
brwWebBrowser.GoBack
End Sub

Private Sub Command2_Click()
brwWebBrowser.GoForward
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen
 brwWebBrowser.Navigate CommonDialog1.FileName
 
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Label1.Caption = File1.FileName
End Sub

Private Sub File1_DblClick()
 brwWebBrowser.Navigate Dir1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
   Label1.Caption = " "
End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub picAddress_Click()

End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub
Public Sub PonerSystray()

  'Tamaño de la estructura systray
  sysTray.cbSize = Len(sysTray)
  'Establecemos el Hwnd, en este caso del formulario
  sysTray.hwnd = UserControl.hwnd
  sysTray.uId = vbNull
  'Flags
  sysTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  'Establecemos el mensaje callback
  sysTray.ucallbackMessage = WM_MOUSEMOVE
  'establecemos el icono, en este caso el que tiene el control Image1
  sysTray.hIcon = Image1.Picture
  'Establecemos el tooltiptext
  sysTray.szTip = m_ToolTiptext & vbNullChar
  'Ponemos el icono en el systray
  Shell_NotifyIcon NIM_ADD, sysTray

End Sub

