VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VeraSoft Develoment"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "frmRegEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************
'Referencias para que Funcione este Program
'Ir a menu Project y despues referencias y agrega Windows Script Host Object Model

'*************************************************************
' Constante de la rama del registro donde estan los paths d las
'aplicaciones
Const Rama_Windows_Run As String = "HKEY_LOCAL_MACHINE\SOFTWARE\" & _
"Microsoft\Windows\CurrentVersion\Run\"
'*************************************************************
'Variable de objeto para poder usar Windows Script Host
Dim o_Registro As WshShell
'*************************************************************

Private Sub Command1_Click()
'*************************************************************
' Grabamos la clave con el metodo RegWrite
' la funcion App es para obtener propiedades del programa
'*************************************************************

Call o_Registro.RegWrite(Rama_Windows_Run & App.EXEName, App.Path & "\" _
& App.EXEName & ".exe")
Call habilitar_botones
End Sub

Private Sub Command2_Click()
'*************************************************************
' Borramos la clave con el metodo RegDelete
Call o_Registro.RegDelete(Rama_Windows_Run & App.EXEName)
habilitar_botones
End Sub

Private Sub Form_Load()
'*************************************************************
'creamos e insanciamos una variale`para poder usar las funciones de windows SH
Set o_Registro = New WshShell
Command1.Caption = "Grabar Ruta"
Command2.Caption = "Eliminar Ruta"
Command2.Enabled = False

End Sub
Sub habilitar_botones()
Command2.Enabled = Not Command2.Enabled
Command1.Enabled = Not Command1.Enabled
End Sub
