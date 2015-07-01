VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USB VERSION 1.0 VERA SOFTDESING"
   ClientHeight    =   6990
   ClientLeft      =   2355
   ClientTop       =   3555
   ClientWidth     =   12675
   ForeColor       =   &H8000000F&
   Icon            =   "Aplicacion Usb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   12675
   Begin VB.OptionButton optUnidad 
      BackColor       =   &H80000007&
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
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   2895
   End
   Begin VB.OptionButton optCarpeta 
      BackColor       =   &H80000007&
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
      Left            =   120
      TabIndex        =   5
      Top             =   5400
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
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.DriveListBox drvPrueba 
      BackColor       =   &H80000006&
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
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.FileListBox filPrueba 
      BackColor       =   &H80000001&
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
      Height          =   4380
      Left            =   3360
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.DirListBox dirPrueba 
      BackColor       =   &H80000006&
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
      Height          =   3915
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
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
      Height          =   495
      Left            =   3120
      MaskColor       =   &H00808080&
      Picture         =   "Aplicacion Usb.frx":3150
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DirPath As String
Private Sub Command1_Click(Index As Integer)
Programador.Show

End Sub

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

Private Sub Form_Load()
filPrueba.Pattern = "*.JPEG;*.bmp;*.jpg"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optCarpeta_Click()
Form3.Show
optCarpeta = False
End Sub

Private Sub optUnidad_Click()
Form2.Show
optUnidad = False

End Sub
