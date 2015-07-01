VERSION 5.00
Object = "*\ASysTray.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin sysTray.Tray Tray1 
      Left            =   7080
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      IconPicture     =   "Form1.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "Quitar del Tray"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Colocar en Tray"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cambiar Icono"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cambiar ToolTipText"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
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
'Eliminar el Tray
Tray1.RemoverSystray
End Sub


'Cambia el ícono del systray
Private Sub Command4_Click()
With CommonDialog1
     .Filter = "Archivos de iconos|*.ico"
     .ShowOpen
     If .FileName = "" Then Exit Sub
     Tray1.IconPicture = LoadPicture(.FileName)
End With

End Sub


'Cambia el toolTipText
Private Sub Command5_Click()

Tray1.ToolTipText = Text1

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Eliminar el Tray
Tray1.RemoverSystray
End Sub

Private Sub Tray1_DblClick(Button As Integer)
'Iquierdo
If Button = vbLeftButton Then
   Me.Show
   Tray1.RemoverSystray
End If
End Sub


