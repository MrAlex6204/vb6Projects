VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELECCIONE LA UNIDAD"
   ClientHeight    =   3795
   ClientLeft      =   4740
   ClientTop       =   5550
   ClientWidth     =   7605
   DrawMode        =   1  'Blackness
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Usb.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "Usb.frx":2832
      Left            =   1200
      List            =   "Usb.frx":2877
      TabIndex        =   3
      Text            =   "Color Texto"
      Top             =   2280
      Width           =   4695
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
      Left            =   2520
      MaskColor       =   &H000000C0&
      Picture         =   "Usb.frx":292E
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.DriveListBox drvDiks 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
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
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Seleccione la Unidad  para aplicar el fondo "
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
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Form1.DirPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
FileCopy Form1.filPrueba.Path + "\" + Form1.filPrueba.FileName, drvDiks.Drive + "\FondoUsb.jpg"
Open drvDiks.Drive + "\DESKTOP.INI" For Output As #1 'genera el archivo el el drive
'seleccionado por drvDisk

Print #1, "[{BE098140-A513-11D0-A3A4-00C04FD706EC}]"
Print #1, "ICONAREA_IMAGE=FondoUsb.jpg"
Print #1, "ICONAREA_TEXT=" + Color_Texto.Text

Close #1
MsgBox "Fondo aplicado Porfavor Actualiza la Unidad Para Ver los Cambios", vbExclamation_, "Fondo Aplicado"
End If

Form2.Hide
Form1.Show



End Sub

