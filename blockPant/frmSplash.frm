VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   10830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   600
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5130
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   8520
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   4560
         TabIndex        =   5
         Top             =   3960
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3960
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3480
         Width           =   2325
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   4560
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   3480
         Width           =   1245
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VeraSoft Develoment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   555
         Left            =   2760
         TabIndex        =   1
         Top             =   2280
         Width           =   4830
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'*************************************************************

Dim Bloqueado As Boolean
' variable para establecer los segundos de bloqueo
Dim TiempoBloqueo As Integer

'*************************************************************
'Referencias para que Funcione este Program
'Ir a menu Project y despues referencias y agrega Windows Script Host Object Model

'*************************************************************
' Constante de la rama del registro donde estan los paths d las
'aplicaciones
Const Rama_Windows_Run As String = "HKEY_LOCAL_MACHINE\SOFTWARE\" & _
"Microsoft\Windows\CurrentVersion\Run\"
'*************************************************************
'Variable de objeto para poder usar Windows Script Host
Dim o_Registro As WshShell
'*************************************************************

Dim DRIVES(1 To 26) As String
Dim prgName As String







Private Sub cmdOK_Click()

 If txtPassword = "password" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
       Call ProgressBar1_Click
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        'Bloqueado = False _
        Bloqueado = True _
        Bloquear
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox "teclas blok"
Beep

End Sub



Private Sub Form_Load()
 

'*************************************************************
'creamos e insanciamos una variale`para poder usar las funciones de windows SH



Transparent.Aplicar_Transparencia frmSplash.hwnd, 215
'------------------------------
'para poner el form en pantalla completa
' al atamaño de la resolucion  del monitor
frmSplash.Top = 0
frmSplash.Left = 0
frmSplash.Width = Screen.Width
frmSplash.Height = Screen.Height
Frame1.Top = Screen.Height / 4
Frame1.Left = Screen.Width / 4
'-------------------------------


'blckea el muse y el teclad

   
    
    
End Sub



Private Sub lblCopyright_Click()
End Sub

Sub ProgressBar1_Click()
ProgressBar1.Visible = True
Dim i, x As Integer
i = 1
ProgressBar1.Max = 100000


While i < 100000

i = i + 1
ProgressBar1.Value = i

Wend
Unload Me

End Sub





'sub que elimina los Hook para el teclado y mouse
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Desbloquear()
    
    ' Vuelve a Habilitar el teclado
    If IdKeyBoard <> 0 Then UnhookWindowsHookEx IdKeyBoard
    ' Vuelve a Habilitar el mouse
    If IdMouse <> 0 Then UnhookWindowsHookEx IdMouse
    
    ' cambia el flag
    Bloqueado = False
    
    ' cierra el timer y restaura la ventana
    'Timer1.Enabled = False
    Me.WindowState = vbNormal
    Me.Cls
End Sub

Private Sub Bloquear()
    
    Me.WindowState = vbMaximized
    'Timer1.Enabled = True
    
    ' Pone la ventana Always OnT op
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    
    ' deshabilita el teclado
    IdKeyBoard = SetWindowsHookEx(WH_KEYBOARD_LL, _
                                      AddressOf WinProcKeyBoard, _
                                      App.hInstance, 0)
    
    ' deshabilita el mouse
    IdMouse = SetWindowsHookEx(WH_MOUSE_LL, _
                                        AddressOf WinProcMouse, _
                                        App.hInstance, 0)
    ' setea el flag
    Bloqueado = True
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub Timer3_Timer()
Dim x As Long
    x = SystemParametersInfo(97&, True, False, 0&)
End Sub



