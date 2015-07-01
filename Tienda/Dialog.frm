VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Articulos"
   ClientHeight    =   3180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset  'Objeto para manejar la Tabla
Dim SQL As String
Private Sub Form_Load()

Set BDD = OpenDatabase(App.Path & "\Tienda.mdb")

SQL = "SELECT * FROM ListArt"

Set TBL = BDD.OpenRecordset(SQL)   'TBL almacena todos los valores de la tabla



TBL.MoveFirst  'nos posicionamos en el primer registro de la tabla


Do Until TBL.EOF  ''La propiedad EOF se pone TRUE cuando se a llegado al final de la tabla
   List1.AddItem TBL("NumArt") & "-->" & TBL("Descrip") & "-->" & TBL("Precio")
   TBL.MoveNext   'pasamos al siguiente registro
Loop


End Sub

Private Sub OKButton_Click()
Me.Hide

End Sub
