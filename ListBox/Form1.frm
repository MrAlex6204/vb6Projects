VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Ingresar"
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Añadir"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
   MsgBox "Debe ingresar un nombre para poder agregar un elemento", vbQuestion + vbOKOnly, "Datos incompletos"
   'Salimos de la rutina ya que no se ha ingresado nada en el control text1
   Exit Sub
End If

'Agregamos el contenido del Text1 en el control List1
List1.AddItem Text1
End Sub

Private Sub Command2_Click()
'Si la lista no está vacía entonces podemos eliminar
If List1.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub List1_Click()

If List1.ListIndex <> -1 Then
   'para pbtener el dato de un ittem seleccionado
  Label2 = List1.List(List1.ListIndex)

End If

End Sub
