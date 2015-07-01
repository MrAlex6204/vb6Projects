VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10740
   LinkTopic       =   "Form3"
   ScaleHeight     =   7380
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
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
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   5760
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "&Acepar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      MaskColor       =   &H000000C0&
      Picture         =   "Form3.frx":0000
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox Color_Texto 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   420
      ItemData        =   "Form3.frx":17B2
      Left            =   4440
      List            =   "Form3.frx":17F7
      TabIndex        =   1
      Text            =   "Color Texto"
      Top             =   2280
      Width           =   4695
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
      Height          =   3690
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Path de la Carpeta Seleccionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Seleccione el Texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Form1.DirPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
FileCopy Form1.filPrueba.Path + "\" + Form1.filPrueba.FileName, dirPrueba.Path + "\FondoUsb.bmp"
Open dirPrueba.Path + "\DESKTOP.INI" For Output As #1  'genera el archivo el el drive
'seleccionado por drvDisk

Print #1, "[{BE098140-A513-11D0-A3A4-00C04FD706EC}]"
Print #1, "ICONAREA_IMAGE=FondoUsb.bmp"
Print #1, "ICONAREA_TEXT=" + Color_Texto.Text

Close #1
MsgBox "Fondo aplicado Porfavor Actualiza la Unidad Para Ver los Cambios", vbExclamation_, "Fondo Aplicado"
End If

Form3.Hide
Form1.Show
End Sub

Private Sub dirPrueba_Change()

dirPrueba.Path = drvPrueba.Drive
Text1.Text = dirPrueba.Path

End Sub

Private Sub drvPrueba_Change()
dirPrueba.Path = drvPrueba.Drive
End Sub

Private Sub Form_Load()
Text1.Text = dirPrueba.Path
End Sub
