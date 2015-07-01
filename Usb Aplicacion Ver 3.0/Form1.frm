VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5175
   ClientLeft      =   4920
   ClientTop       =   3570
   ClientWidth     =   4125
   ControlBox      =   0   'False
   DrawMode        =   10  'Mask Pen
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   206.25
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   360
      MaskColor       =   &H000000C0&
      Picture         =   "Form1.frx":0000
      TabIndex        =   5
      Top             =   4880
      Width           =   495
   End
   Begin VB.DriveListBox drvDiks 
      BackColor       =   &H00000000&
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
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
      Height          =   375
      Left            =   1080
      MaskColor       =   &H000000C0&
      Picture         =   "Form1.frx":17B2
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox Color_Texto 
      BackColor       =   &H00000000&
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
      ItemData        =   "Form1.frx":2F64
      Left            =   360
      List            =   "Form1.frx":2FA9
      TabIndex        =   0
      Text            =   "Color Texto"
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   4320
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
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
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   3240
      Picture         =   "Form1.frx":3060
      Stretch         =   -1  'True
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3360
      Picture         =   "Form1.frx":55BD
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3360
      Picture         =   "Form1.frx":7D55
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   -360
      Picture         =   "Form1.frx":A502
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me


End Sub


 Sub Command3_Click()
Image1.Picture = Image2.Picture



End Sub

 Sub Command4_Click()
Image1.Picture = Image4.Picture


End Sub

 Sub Command5_Click()
Image1.Picture = Image3.Picture


End Sub

      Private Sub Form_Load()
         
         
         Form2.Top = Form1.Top + 30
         Form2.Left = Form1.Left + Form1.Width
          Mover.MoverForm
         
      End Sub



      Private Sub Command1_Click()
        On Error Resume Next
If Form1.Text1 = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
FileCopy Form1.Text1.Text, drvDiks.Drive + "\FondoUsb.jpg"
Open drvDiks.Drive + "\DESKTOP.INI" For Output As #1 'genera el archivo el el drive
'seleccionado por drvDisk

Print #1, "[{BE098140-A513-11D0-A3A4-00C04FD706EC}]"
Print #1, "ICONAREA_IMAGE=FondoUsb.jpg"
Print #1, "ICONAREA_TEXT=" + Color_Texto.Text

Close #1
MsgBox "Fondo aplicado Porfavor Actualiza la Unidad Para Ver los Cambios", vbExclamation_, "Fondo Aplicado"
End If
        
        
        
      End Sub





Private Sub Timer1_Timer()
Mover.MoverForm
End Sub
