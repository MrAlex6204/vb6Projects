VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2700
   ClientLeft      =   2100
   ClientTop       =   3975
   ClientWidth     =   4695
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "  Nombre del USD o Disco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Etiqueta 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Tag             =   "EL nombre de la unidad no debe llevar Caracteres especiales "
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.DriveListBox drvDiks 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la Unidad  para aplicar AutoRun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label LabelAplicar 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   720
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
If Check1.Value = 1 Then
Etiqueta.Visible = True
Else
Etiqueta.Visible = False
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabelAplicar.FontSize = 10
LabelAplicar.FontBold = False
LabelAplicar.ForeColor = &H808000

Dim lngReturnValue As Long


        If Button = 1 Then
        
         Mover.MoverForm Form3.hWnd
     
        End If
End Sub

Private Sub LabelAplicar_Click()

If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
'
FileCopy Form1.dialogo.FileName, drvDiks.Drive + "\" + Form1.dialogo.FileTitle

Open drvDiks.Drive + "\AUTORUN.INF" For Output As #1 'genera el archivo el el drive
'seleccionado por drvDisk


Print #1, "[autorun]"

If Check1.Value = 1 Then
    Print #1, "action=Run " + Etiqueta
End If

Print #1, "Icon=" + Form1.dialogo.FileTitle

If Check1.Value = 1 Then
    Print #1, "Label=" + Etiqueta
End If

Close #1


MsgBox "Icono aplicado Porfavor Actualiza la Unidad Para Ver los Cambios", vbExclamation_, "Fondo Aplicado"
End If
End Sub

Private Sub LabelAplicar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.RepSound App.Path + "\Sounds\Sound2.wav"
Form1.Reproduce = True

LabelAplicar.FontSize = 12
LabelAplicar.FontBold = True
LabelAplicar.ForeColor = &HC0C000
End Sub

