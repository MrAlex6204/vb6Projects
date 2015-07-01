VERSION 5.00
Begin VB.Form frmGuardar 
   Caption         =   "Guaradar"
   ClientHeight    =   2685
   ClientLeft      =   4590
   ClientTop       =   4380
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4965
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFilenom 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nom. Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmGuardar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
Open "Temp.txt" For Append As #1

'NOTA:
'si no se especifica la direecion de almacenamiento del fichero
'el fichero se crea dentro del mismo irectorio del programa

Print #1, txtFilenom.Text, frmTemp.vsbTemp.value
Close #1


txtFilenom.Text = Empty
frmGuardar.Hide

End Sub

