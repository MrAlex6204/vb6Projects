VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   4350
   ClientTop       =   3225
   ClientWidth     =   7215
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar Progress 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtDesencrip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtEncrip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblMinombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5160
      TabIndex        =   2
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3450
      Left            =   4440
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2730
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   6720
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblBar 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Label Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "VERASOFT DEVELOPMENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   6885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encriptacion Metodo Cesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Label lblResul2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label Desencriptar 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Desencriptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   2325
   End
   Begin VB.Label lblEncrip 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encriptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1635
   End
   Begin VB.Label lblResul 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape ShapeBorder 
      BorderWidth     =   5
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Shape ShapeBar 
      BorderWidth     =   3
      Height          =   735
      Left            =   240
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Desencriptar_Click()
If txtDesencrip = Empty Then
Beep
Exit Sub
End If
Progress.Visible = True
Dim n, I, j As Integer
Dim Caracter, text, LETRA As String
n = Len(txtDesencrip)
Progress.Max = n
I = 1
Do While (n >= I)
Progress.Value = I
 j = 0
Caracter = Mid(UCase(txtDesencrip), I, 1)
   
    Do While (28 > j)
        If DESENCRIP(j) = Caracter Then
        LETRA = ENCRIP(j)
        text = text + LETRA
        Exit Do
        End If
        j = j + 1
    Loop
    


I = I + 1

Loop


lblResul2 = text
End Sub

Private Sub Desencriptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Desencriptar.FontSize = 24
Desencriptar.ForeColor = &HC0&
End Sub

Private Sub Form_Load()

Call CargarSonido("C:\Sonido.mp3") ' Llama a la funcion q crea el archivo temporal
Call mciExecute("play " & "C:\Sonido.mp3") ' Reproduce el archivo temporal
lblResul = Empty
lblResul2 = Empty
Call ENCRIPDAT
Call DESENCRIPDAT

'Configuracion de Fuente de los Controles
lblTitle.Font = "MadScience"
lblTitle.FontSize = 24

lblEncrip.Font = "MadScience"
Desencriptar.Font = "MadScience"

Label1.Font = "MadScience"
Label1.FontSize = 24
txtEncrip.Font = "MadScience"
txtDesencrip.Font = "MadScience"
lblClose.Font = "MadScience"
lblMinombre.Font = "MadScience"
lblBar = Empty


Form1.Refresh


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEncrip.FontSize = 20
lblEncrip.ForeColor = &H0&

lblClose.FontBold = False
lblClose.ForeColor = &H0&
lblClose.FontSize = 20

Desencriptar.FontSize = 20
Desencriptar.ForeColor = &H0&
ShapeBar.BorderColor = &H0&


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
mciExecute "Close All" 'Detiene la reproduccion
Kill App.path + "\sonido.Mp3" ' elimina el archivo temporal
End Sub

Private Sub lblBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Transparent.Aplicar_Transparencia Me.hWnd, 150
        
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
        Mover.MoverForm
End If
Transparent.Aplicar_Transparencia Me.hWnd, 255

lblClose.FontBold = False
lblClose.ForeColor = &H0&
lblClose.FontSize = 20

ShapeBar.BorderColor = &HC0&

End Sub

Private Sub lblClose_Click()
mciExecute "Close All" 'Detiene la reproduccion
Kill "C:\Sonido.Mp3" ' elimina el archivo temporal
Kill App.path + "\MadSience.ttf"
End
End Sub


Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.FontBold = True
lblClose.FontSize = 24
lblClose.ForeColor = &HC0&
End Sub

Private Sub lblEncrip_Click()

If txtEncrip = Empty Then
Beep
Exit Sub
End If


Progress.Visible = True
Dim n, I, j As Integer
Dim Caracter, text, LETRA As String
n = Len(txtEncrip)
Progress.Max = n
I = 1
Do While (n >= I)
Progress.Value = I
 j = 0
Caracter = Mid(UCase(txtEncrip), I, 1)
   
    Do While (28 > j)
   
        If ENCRIP(j) = Caracter Then
        LETRA = DESENCRIP(j)
        text = text + LETRA
        Exit Do
        End If
        j = j + 1
    Loop
    


I = I + 1

Loop


lblResul = text
End Sub

Private Sub lblEncrip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEncrip.FontSize = 24
lblEncrip.ForeColor = &HC0&
End Sub

Private Sub txtDesencrip_Change()
 lblResul2 = Empty
Progress.Min = 0
Progress.Value = 0
Progress.Visible = False
End Sub

 Sub txtEncrip_Change()
lblResul = Empty
Progress.Min = 0
Progress.Value = 0
Progress.Visible = False
End Sub

