VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USB VERSION 1.0 VERA SOFTDESING"
   ClientHeight    =   6225
   ClientLeft      =   2355
   ClientTop       =   3555
   ClientWidth     =   12675
   ForeColor       =   &H8000000F&
   Icon            =   "Aplicacion Usb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   12675
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "Solo Inmagenes"
      DialogTitle     =   "Usb Aplication VeraSoft Develoment"
      Filter          =   "Imágenes(*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Examinar"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3855
   End
   Begin VB.OptionButton optUnidad 
      BackColor       =   &H00000000&
      Caption         =   "Aplicar Fondo a &Disco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.OptionButton optCarpeta 
      BackColor       =   &H00000000&
      Caption         =   "Aplicar Fondo a &Carpeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdCreador 
      Caption         =   "&Creador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H000080FF&
      Caption         =   "&Salir"
      DisabledPicture =   "Aplicacion Usb.frx":2832
      DownPicture     =   "Aplicacion Usb.frx":2B3C
      DragIcon        =   "Aplicacion Usb.frx":2E46
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00808080&
      Picture         =   "Aplicacion Usb.frx":3150
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DirPath As String
Private Sub Dir1_Change()

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Form2.Show

End Sub

Private Sub cmdCreador_Click()
Programador.Show
End Sub

Private Sub cmdSalir_Click()
End
End Sub

Private Sub dirPrueba_Change()
filPrueba.Path = dirPrueba.Path
'filPrueba.Path obtiene la direccion del docuemnto a file
'y filPrueba.FileName obtiene el nombre del docuemto


End Sub

Private Sub drvPrueba_Change()

dirPrueba.Path = drvPrueba.Drive
End Sub

Private Sub filPrueba_Click()
DirPath = filPrueba.Path + "\" + filPrueba.FileName

Image1.Picture = LoadPicture(filPrueba.Path + "\" + filPrueba.FileName)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Text1.Text = CommonDialog1.FileName
'Cargamos la imagen del path que tiene text1

End If
'Image1.Picture = LoadPicture(Text1.Tex)
End Sub

Private Sub optCarpeta_Click()
Form3.Show
optCarpeta = False
End Sub

Private Sub optUnidad_Click()
Form2.Show
optUnidad = False

End Sub

Private Sub Text1_Change()
Image1.Picture = LoadPicture(Text1)
End Sub
