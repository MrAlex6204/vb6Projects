VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00080000&
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   3240
      TabIndex        =   3
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   2040
      Picture         =   "frmTransparente.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
'valores maximos y minimos
HScroll1.Max = 255
HScroll1.Min = 20

'le establecemos un valor por defecto a la barra
HScroll1.Value = 150
Label1.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Change()
'llamamos la funcio pasandole el handle del form y el valor de la _
transparencia , que es el de la barra
Beep
'******************************************************************* _
Madamos a llamar ala Funcion Aplicar_Tranparencia del Modulo _
Transparent para que aplique el efecto de Transparencia
Transparent.Aplicar_Transparencia Form1.hwnd, CByte(HScroll1.Value)
Label1.Caption = HScroll1.Value
'*******************************************************************

End Sub
