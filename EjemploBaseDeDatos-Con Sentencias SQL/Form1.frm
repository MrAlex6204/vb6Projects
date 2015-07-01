VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "&ELIMINAR"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "VER REGISTROS"
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AGR. NUEVO"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIM. DUPLICADOS"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MAX MIN"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "WHERE"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GROUP BY"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Referencias del Proyecto para Ejecutar es Proyecto _
Referencia y seleccionamos Microsoft DAO 3.51 Object


Dim SQL As String


Private Sub Command1_Click()
List1.Clear
'Group By _
Esta cláusula se utiliza para agrupar segun lo q _
especifiquemos. Por ejemplo, podemos listar todos los datos _
de nuestra tabla1, pero agrupados por edad.

Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos

SQL = "SELECT edad, nombre, apellido FROM tabla1 GROUP BY edad,nombre, apellido"
Set TBL = BDD.OpenRecordset(SQL)

TBL.MoveFirst
Do Until TBL.EOF
List1.AddItem TBL("nombre") & " " & TBL("apellido") & " tiene " & TBL("edad")
TBL.MoveNext
Loop

TBL.Close
BDD.Close


End Sub

Private Sub Command2_Click()
List1.Clear

Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos

SQL = "SELECT * FROM tabla1 WHERE edad < 30"
Set TBL = BDD.OpenRecordset(SQL)

TBL.MoveFirst
Do Until TBL.EOF
List1.AddItem TBL("edad")
TBL.MoveNext
Loop
TBL.Close
BDD.Close


End Sub

Private Sub Command3_Click()
List1.Clear

Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos

SQL = "SELECT COUNT(*), MIN(edad), AVG(edad) FROM tabla1"
Set TBL = BDD.OpenRecordset(SQL)

List1.AddItem "total de reg: " & TBL("expr1000")
List1.AddItem "MINIMA EDAD: " & TBL("expr1001")
List1.AddItem "PROMEDIO EDADES: " & TBL("expr1002")

TBL.Close
BDD.Close


End Sub

Private Sub Command4_Click()
List1.Clear

Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos

SQL = "SELECT DISTINCT edad FROM tabla1"  'almacena todas las edades sin repetirlas
Set TBL = BDD.OpenRecordset(SQL)


TBL.MoveFirst
Do Until TBL.EOF
   List1.AddItem TBL("edad")
   TBL.MoveNext
Loop

TBL.Close
BDD.Close


End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
List1.Clear

Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset 'Objeto para manejar la Tabla
Set BDD = OpenDatabase(App.Path + "\BaseDeDatos\BaseDeDatos.mdb") 'Abre la base de datos

SQL = "SELECT * FROM tabla1"
Set TBL = BDD.OpenRecordset(SQL)



TBL.MoveFirst
    Do Until TBL.EOF
    List1.AddItem TBL("Nombre") & " " & TBL("Apellido") & " Tiene--> " & TBL("Edad")
    TBL.MoveNext
Loop

TBL.Close
BDD.Close


End Sub

Private Sub Command7_Click()
Form3.Show

End Sub

Private Sub Label1_Click()

End Sub

