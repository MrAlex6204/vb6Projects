VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form VeraSoft 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "VeraSoft Develooment"
   ClientHeight    =   7230
   ClientLeft      =   3975
   ClientTop       =   3030
   ClientWidth     =   8535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Interval        =   85
      Left            =   3840
      Top             =   2280
   End
   Begin ComctlLib.ProgressBar Progress 
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame FrameMensaje 
      BackColor       =   &H80000009&
      Height          =   1455
      Left            =   1920
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.Label CloseMensaje 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Punk Kid"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   75
         TabIndex        =   18
         Top             =   720
         Width           =   4380
      End
   End
   Begin VB.TextBox txtClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtDesencrip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox txtEncrip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblMarquesina 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "MARQUESINA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   8085
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   7395
      TabIndex        =   14
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblMiNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6480
      TabIndex        =   13
      Top             =   4920
      Width           =   450
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   7935
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Label bar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STATUSBAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   8055
   End
   Begin VB.Shape ShapeBar 
      BorderWidth     =   4
      Height          =   735
      Left            =   240
      Top             =   120
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BorderWidth     =   8
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4290
      Left            =   5040
      Picture         =   "Form1.frx":08CA
      Top             =   1920
      Width           =   3330
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "verasoft development"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   7980
   End
   Begin VB.Label lblNuevaClave 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVA  CLAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblGenerar 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblResul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
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
      TabIndex        =   7
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Label lblEncrip 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Encriptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Desencriptar 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Desencriptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1470
   End
   Begin VB.Label lblResul2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
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
      TabIndex        =   4
      Top             =   5040
      Width           =   1515
   End
End
Attribute VB_Name = "VeraSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strText As String
Private Sub bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblClose.FontBold = False
lblClose.ForeColor = &H0&
lblClose.FontSize = 18

lblMin.FontBold = False
lblMin.ForeColor = &H0&
lblMin.FontSize = 18

ShapeBar.BorderColor = &HC0&

If Button = 1 Then
Transparent.Aplicar_Transparencia Me.hWnd, 150
        
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
     '   Mover.MoverForm
End If

Transparent.Aplicar_Transparencia Me.hWnd, 255

End Sub

Private Sub Command1_Click()
txtClave.Enabled = True
lblGenerar.Visible = True
Kill App.path + "\Encript2.Dat"
Kill App.path + "\Encript.Dat"

txtEncrip.Enabled = False
txtDesencrip.Enabled = False
Command1.Visible = False
End Sub

Private Sub CloseMensaje_Click()
FrameMensaje.Visible = False

End Sub

Private Sub CloseMensaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseMensaje.ForeColor = &HC0&
CloseMensaje.FontSize = 24
End Sub

Private Sub DesEncriptar_Click()
If txtDesencrip = Empty Then
Declaraciones.Mensaje ("Introdusca Mas de 1 Caracteres")
Exit Sub
End If
'Antes Receteamos la Variable
DesencriptacionText = Empty
Call Desencriptardor(txtDesencrip)
lblResul2.Visible = True
lblResul2 = DesencriptacionText
lblStat.Visible = True
lblStat = "Desencriptacion Generada"
End Sub

Private Sub Desencriptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Desencriptar.FontSize = 20
Desencriptar.ForeColor = &HC0&
End Sub

Private Sub Form_Load()
Call CargarSonido("C:\Sonido.Mp3") ' Llama a la funcion q crea el archivo temporal
Call mciExecute("play " & "C:\Sonido.Mp3") ' Reproduce el archivo temporal
lblResul = Empty
lblResul2 = Empty
Call AlfabetoLoad
bar = Empty
lblMin.Font = "punk kid"
lbl2.Font = "punk kid"
lblStat.Font = "punk kid"
lblMiNombre.Font = "punk kid"
lblMarquesina.Font = "punk kid"
Label3.Font = "punk kid"
lblGenerar.Font = "punk kid"
lblNuevaClave.Font = "punk kid"
lblEncrip.Font = "punk kid"
Desencriptar.Font = "punk kid"
lblClose.Font = "punk kid"
txtClave.Font = "punk kid"
txtEncrip.Font = "punk kid"
txtDesencrip.Font = "punk kid"
lblMensaje.Font = "punk kid"
lblMensaje.FontSize = 12

strText = String(50, " ") + "U.A.T   I.S.C  Metodo MonoAlfabetico"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ShapeBar.BorderColor = &H0&

lblMarquesina.Font = "punk kid"
Label3.Font = "punk kid"
lblGenerar.Font = "punk kid"
lblNuevaClave.Font = "punk kid"
lblEncrip.Font = "punk kid"
Desencriptar.Font = "punk kid"
lblClose.Font = "punk kid"
txtClave.Font = "punk kid"
txtEncrip.Font = "punk kid"
txtDesencrip.Font = "punk kid"


lblEncrip.FontSize = 12
lblEncrip.ForeColor = &H0&

lblClose.FontBold = False
lblClose.ForeColor = &H0&
lblClose.FontSize = 18

lblMin.FontBold = False
lblMin.ForeColor = &H0&
lblMin.FontSize = 18


Desencriptar.FontSize = 12
Desencriptar.ForeColor = &H0&

lblGenerar.FontSize = 12
lblGenerar.ForeColor = &H0&

lblNuevaClave.FontSize = 12
lblNuevaClave.ForeColor = &H0&

End Sub

Private Sub Form_Unload(Cancel As Integer)
mciExecute "Close All" 'Detiene la reproduccion
Kill "c:\sonido.Mp3" ' elimina el archivo temporal
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label2_Click()


End Sub

Private Sub FrameMensaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CloseMensaje.FontBold = False
CloseMensaje.ForeColor = &H0&
CloseMensaje.FontSize = 18
End Sub

Private Sub lblClose_Click()
On Error Resume Next
mciExecute "Close All" 'Detiene la reproduccion
Kill "c:\sonido.Mp3" ' elimina el archivo temporal
Kill App.path + "\Encript2.Dat"
Kill App.path + "\Encript.Dat"
Kill "C:\punk kid.ttf"
End
End Sub


Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &HC0&
lblClose.FontSize = 24
End Sub

Private Sub lblEncrip_Click()



If txtEncrip = Empty Then
Declaraciones.Mensaje ("Introdusca Mas de 1 Caracteres")
Exit Sub
End If

'Antes Receteamos la Variable
EncriptacionText = Empty
Call Encriptar(txtEncrip)
lblResul.Visible = True
lblResul = EncriptacionText
lblStat.Visible = True
lblStat = "Encriptacion Generada"
End Sub

Private Sub lblEncrip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEncrip.FontSize = 20
lblEncrip.ForeColor = &HC0&
End Sub

Private Sub lblGenerar_Click()
Dim Num As Integer
Num = Len(txtClave)
If txtClave = Empty Or Num <= 3 Then

Declaraciones.Mensaje ("Introdusca Mas de 3 Caracteres")

Exit Sub
End If

Dim Contenido As String
Generar (txtClave)
GenerarEncriptacioFiles (txtClave)

txtClave.Enabled = False

lblStat.Visible = True
lblStat = "Archivos Generados"
lblGenerar.Visible = False
txtEncrip.Enabled = True
txtDesencrip.Enabled = True

lblNuevaClave.Visible = True

End Sub

Private Sub lblGenerar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGenerar.FontSize = 20
lblGenerar.ForeColor = &HC0&
End Sub

Private Sub lblMin_Click()
Me.WindowState = 1
End Sub

Private Sub lblMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMin.ForeColor = &HC0&
lblMin.FontSize = 24
End Sub

Private Sub lblNuevaClave_Click()

txtClave.Enabled = True
lblGenerar.Visible = True
Kill App.path + "\Encript2.Dat"
Kill App.path + "\Encript.Dat"

txtEncrip.Enabled = False
txtDesencrip.Enabled = False
lblNuevaClave.Visible = False
End Sub

Private Sub lblNuevaClave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNuevaClave.FontSize = 20
lblNuevaClave.ForeColor = &HC0&
End Sub

Private Sub Timer_Timer()
strText = Mid(strText, 2) & Left(strText, 1)
lblMarquesina = strText
End Sub

Private Sub txtDesencrip_Change()
Progress.Visible = False
lblStat.Visible = False
End Sub

Private Sub txtDesencrip_KeyPress(KeyAscii As Integer)
If IsNumeric(txtDesencrip) = True Then
Beep
Declaraciones.Mensaje ("No se Aceptan Numeros!!!")
Exit Sub
End If
End Sub

Sub txtEncrip_Change()
Progress.Visible = False
lblStat.Visible = False
End Sub

Private Sub txtEncrip_KeyPress(KeyAscii As Integer)
If IsNumeric(txtEncrip) = True Then
Beep
Declaraciones.Mensaje ("No se Aceptan Numeros!!!")
Exit Sub
End If
End Sub
