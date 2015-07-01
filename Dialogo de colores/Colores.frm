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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

    Dim El_Color As Long
    'La variable El_Color almacenará el color en formato Long
    'del color elegido. Si no se eligió ninguno retornamos desde
    'la función el valor -1, si no establecemos el color defondo
    'del form pasandole el valor devuelto por la función
    
    ' llamamos al cuadro diálogo Seleccionar Color
    El_Color = Abrir_CommonDialog_Color(Me)
    
    If El_Color <> -1 Then
        ' establecemos el color de fondo del Form con el color seleccionado
        Label1.BackColor = El_Color
        Label1 = El_Color
    Else
        MsgBox "Se canceló ", vbInformation
    End If
End Sub

Private Sub Form_Load()
Command1.Caption = " Seleccionar Color "
End Sub

