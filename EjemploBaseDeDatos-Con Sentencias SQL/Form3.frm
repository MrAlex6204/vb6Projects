VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&<<"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&>>"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Apellido"
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
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Edad:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   630
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla


Private Sub Command1_Click()
On Error GoTo salir
TBL.MovePrevious

If TBL.EOF Then   ''EOF esta en verdaero si no hay datos
   TBL.MoveLast
   Exit Sub
Else

Text1 = TBL("Nombre")
   Text2 = TBL("Apellido")
   Text3 = TBL("Edad")
End If

Exit Sub
salir:
   
TBL.MoveLast
End Sub


Private Sub Command7_Click()

'Set TBL = BDD.OpenRecordset(SQL)
TBL.MoveNext

If TBL.EOF Then  ''EOF esta en verdaero si no hay datos
   TBL.MoveFirst
   Exit Sub
Else
Text1 = TBL("Nombre")
   Text2 = TBL("Apellido")
   Text3 = TBL("Edad")
End If




End Sub

Private Sub Form_Load()


Call setdat


'


TBL.MoveFirst
   Text1 = TBL("Nombre")
   Text2 = TBL("Apellido")
   Text3 = TBL("Edad")


End Sub


Sub setdat()

Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos
SQL = "SELECT * FROM tabla1"
Set TBL = BDD.OpenRecordset(SQL)


End Sub
