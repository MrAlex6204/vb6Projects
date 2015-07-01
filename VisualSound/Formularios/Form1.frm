VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   930
   ClientLeft      =   -30
   ClientTop       =   -285
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar Bar 
      Height          =   375
      Left            =   923
      TabIndex        =   9
      Top             =   278
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblPorcentaje 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2685
      TabIndex        =   10
      Top             =   720
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   105
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   4800
      Picture         =   "Form1.frx":1CFA
      Top             =   105
      Width           =   720
   End
   Begin VB.Label lblPLay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Play"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label lblNext 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
      Height          =   255
      Left            =   2700
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblMedia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Media"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblPrev 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prev"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblStop 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label lblDown 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vol. Down"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label lblMute 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mute"
      Height          =   255
      Left            =   1500
      TabIndex        =   1
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label lblUp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vol. Up"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSVista 
         Caption         =   "Siempre a la Vista"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIWindows 
         Caption         =   "Iniciar con Windows"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private WithEvents C As cMMKeys
Attribute C.VB_VarHelpID = -1
Private WithEvents f_cSystray As cSystray
Attribute f_cSystray.VB_VarHelpID = -1
Dim i As Integer

Private Sub c_KeyEventDown(mmKey As emmKey)
    ChangeState mmKey, True
End Sub

Private Sub c_KeyEventUp(mmKey As emmKey)
    ChangeState mmKey, False
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub f_cSystray_MouseUp(Button As Integer)
    f_cSystray.BeforePopup
    PopupMenu mnuMenu
End Sub

Private Sub Form_Initialize()
    InitCommonControls
    On Error Resume Next
    Dim Abrir As String
    Abrir = App.Path + "\config.ini"

    mnuSVista.Checked = LeerIni("Ventana", "Estado", Abrir)
    mnuIWindows.Checked = LeerIni("Inicio", "Estado", Abrir)
End Sub

Private Sub Form_Load()
    Set C = New cMMKeys
    
    Me.Hide
        
    Bar.Min = 0
    Bar.Max = 100

    If Verificar_tarjeta Then
        Call OpenMixer
        Bar.Value = Volumen
    Else
        MsgBox "No se detectó una tarjeta de sonido en el sistema", vbCritical
    End If
    
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2 + 3000
    
    AgregarIcon
End Sub

Sub AgregarIcon()
    Set f_cSystray = New cSystray
    
    With f_cSystray
        .SysTrayIconFromRes "ICON_0"
        .SysTrayToolTip = "Tooltip"
        .SysTrayShow True
    End With
End Sub

Private Sub ChangeState(sKey As emmKey, bState As Boolean)

'    Dim lColor As Long
'
'    If bState = False Then
'        lColor = vbWhite
'    Else
'        lColor = vbBlue
'    End If
    
    Select Case sKey
'        Case mmKey_LAUNCH_MEDIA_SELECT: 'lblMedia.BackColor = lColor
'        Case mmKey_MEDIA_NEXT_TRACK: 'lblNext.BackColor = lColor
'        Case mmKey_MEDIA_PLAY_PAUSE: 'lblPLay.BackColor = lColor
'        Case mmKey_MEDIA_PREV_TRACK: 'lblPrev.BackColor = lColor
'        Case mmKey_MEDIA_STOP: 'lblStop.BackColor = lColor

        Case mmKey_VOLUME_UP: 'lblUp.BackColor = lColor
            Me.Show
            Timer1.Interval = 1000
            VisualizarVol
        Case mmKey_VOLUME_MUTE: 'lblMute.BackColor = lColor
            
            Bar.Value = 100
            VisualizarVol
            'Me.Caption = "vol = Mute"
        Case mmKey_VOLUME_DOWN: 'lblDown.BackColor = lColor
            Me.Show
            Timer1.Interval = 1000
            VisualizarVol
    End Select
End Sub

Sub VisualizarVol()
    'Me.Caption = "Volumen " & Bar.Value & " %"
     
    Call OpenMixer
    Bar.Value = Volumen
    
    lblPorcentaje.Caption = Volumen & "%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call CloseMixer
    Set f_cSystray = Nothing
End Sub

Private Sub mnuCerrar_Click()
    On Error Resume Next
    Call CloseMixer
    Set f_cSystray = Nothing
    Unload Me
End Sub

Private Sub mnuIWindows_Click()
On Error Resume Next
Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")

If mnuIWindows.Checked = True Then
    mnuIWindows.Checked = False
    El_Objeto.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
    GuardarConfiguracion
Else
    mnuIWindows.Checked = True
    El_Objeto.RegDelete ("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName)
    GuardarConfiguracion
End If

Set El_Objeto = Nothing
End Sub

Private Sub mnuSVista_Click()
If mnuSVista.Checked = True Then
    SiempreVisible Form1, False
    mnuSVista.Checked = True
    GuardarConfiguracion
Else
    SiempreVisible Form1, False
    mnuSVista.Checked = True
    GuardarConfiguracion
End If
End Sub

Private Sub GuardarConfiguracion()
    Dim INI As String, i As Integer, datos As String
    INI = App.Path + "\config.ini"
    EscribirINI "Ventana", "Estado", mnuSVista.Checked, INI
    EscribirINI "Inicio", "Estado", mnuIWindows.Checked, INI
End Sub

Private Sub Timer1_Timer()

i = i + 1

If i = 3 Then
    Me.Hide
    Timer1.Interval = 0
    i = 0
End If
End Sub
