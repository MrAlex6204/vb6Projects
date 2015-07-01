VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "VeraSoft Develoment"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "Ejecuta un Exe Con Boton.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&EXAMINAR"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "VeraSoft Develeoment"
      Filter          =   "Programa exe(*.exe)|*.exe"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   810
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo siguiente ' si produce un error se pasa hasta donde

'dice siguiente

If Text1.Text = Empty Then
MsgBox "Telclee el Path de un Exe", vbCritical, "Teclee La direccion Correcta"

siguiente: 'Apartir de aqui se ejecuta si huvo un error en el prg
Text1.SetFocus
Else
Shell Text1.Text, vbNormalFocus
End If


End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub

