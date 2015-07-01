VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cronómetro"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "Cronometro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   1260
      TabIndex        =   1
      Top             =   180
      Width           =   1275
      Begin VB.CommandButton cmdDetener 
         Caption         =   "&Detener"
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   900
      End
      Begin VB.CommandButton cmdIniciar 
         Caption         =   "&Iniciar"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Left            =   420
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "Cronometro.frx":0442
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label lblCronometro 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Long   'Contador.
Dim Tiempo As String  'Tiempo total transcurrido.

Private Sub cmdDetener_Click()
Timer1.Interval = 0
End Sub

Private Sub cmdIniciar_Click()
I = 0 'Inicializar el contador.
Timer1.Interval = 0    'Detener el cronometro
lblCronometro.Caption = ""  'Limpiar la etiqueta
Timer1.Interval = 1    'Iniciar el cronometro
End Sub

Private Sub cmdSalir_Click()
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29 'Centra el formulario completamente

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
I = I + 1
Tiempo = Format(Int(I / 36000) Mod 24, "00") & ":" & _
         Format(Int(I / 600) Mod 60, "00") & ":" & _
         Format(Int(I / 10) Mod 60, "00") & ":" & _
         Format(I Mod 10, "00")
lblCronometro.Caption = Tiempo
End Sub
