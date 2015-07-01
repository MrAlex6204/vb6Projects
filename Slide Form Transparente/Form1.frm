VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Colocar un formulario Form1, un form2 y un módulo bas. _
'''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
' Código fuente del form1
''''''''''''''''''''''''''''''''''''''''''


' Botón que carga y muestra el form
'''''''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
    Call SlideForm(Form2, MOSTRAR, 200, 5)
End Sub

' Botón que oculta y descarga el form
'''''''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
    Call SlideForm(Form2, OCULTAR, 200, 5)
End Sub

Private Sub Form_Load()
    Command1.Caption = " Show "
    Command2.Caption = " Unload "
End Sub
 

