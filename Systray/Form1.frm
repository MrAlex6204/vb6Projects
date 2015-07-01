VERSION 5.00
Object = "{DB3F8F1D-3ADE-4D2C-BA1A-BACA667F0EE4}#1.0#0"; "SysTrayocx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quitar Icon"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Camiar ToolTip"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar tray"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Colocar tray"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin sysTray.Tray Tray1 
      Left            =   1560
      Top             =   120
      _extentx        =   847
      _extenty        =   847
      tooltiptext     =   "Vera"
      iconpicture     =   "Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Notificamos
Tray1.PonerSystray
'Ocultamos el Form
Me.Hide
End Sub

Private Sub Command2_Click()
Tray1.RemoverSystray
End Sub

Private Sub Command3_Click()
Tray1.ToolTipText = Text1.Text
End Sub

Private Sub Command4_Click()
With CommonDialog1
     .Filter = "Archivos de iconos|*.ico"
     .ShowOpen
     If .FileName = "" Then Exit Sub
     'Le cambiamos el ícono
     Tray1.IconPicture = LoadPicture(.FileName)
End With

End Sub

Private Sub Tray1_DblClick(Button As Integer)
'Si hacemos dobleClick con el botón izquierdo restauramos el form y _quitamos el ícono del Tray
If Button = vbLeftButton Then
Me.Show
Tray1.RemoverSystray
End If
End Sub

Private Sub Tray1_MouseDown(Button As Integer)
'Derecho _
If Button = vbRightButton Then _
   MsgBox "Botón derecho" _
End If _

'Iquierdo _
If Button = vbLeftButton Then _
   MsgBox "Botón izquierdo" _
End If _

End Sub


