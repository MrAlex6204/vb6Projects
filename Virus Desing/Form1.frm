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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' flag para saber si está o no bloqueado
Dim Bloqueado As Boolean
' variable para establecer los segundos de bloqueo
Dim TiempoBloqueo As Integer

' Sub que instala los Hook para bloquear el teclado y mouse
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

' Botón que bloquea el teclado y el mouse
'''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
    Bloqueado = True
    Bloquear
End Sub

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
    Bloqueado = False
    ' tiempo de bloqueo 10 segundos
    TiempoBloqueo = 10
    Me.BackColor = vbRed
    Me.FontSize = 20
    Me.ForeColor = vbWhite
    Me.AutoRedraw = True
    Timer1.Interval = 1000
    Timer1.Enabled = False
    Command1.Caption = "Bloquear"
End Sub

Private Sub Timer1_Timer()
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
   
    Me.Print "Mouse y KeyBoard bloquedo. Tiempo : " & _
              segundos & " de : "; TiempoBloqueo & " ..segundos"
    
End Sub



  
