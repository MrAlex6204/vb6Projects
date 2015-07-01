VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Base de Datos Alumnos"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   6150
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   6150
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Alumnos"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataField       =   "Turno"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Carrera"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Matricula"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "TURNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CARRERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MATRICULA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1545
   End
   Begin VB.Menu movimientos 
      Caption         =   "&Movimientos"
      Begin VB.Menu nuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu guardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu buscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu eliminar 
         Caption         =   "&Eliminar"
      End
   End
   Begin VB.Menu reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu alumnos 
         Caption         =   "&Alumnos"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Alumnos_Click()
On Error GoTo ErrorOpen
'El sig. Codigo es para que DataEnvironment1.Alumnos _
se Conecte con la base de datos el path es segun donde se encuentre al aplicacion _
la funcion App.Path devuelve el path donde se encuentra nuestro prog. _
la Base de Datos se Encuentaen una carpeta llamada MiBaseDeDatos _
y la Base de Datos se llama MiBaseDeDatos.mdb


DataEnvironment1.Alumnos.ConnectionString = "DSN=MS Access Database;DBQ=" _
+ App.Path + "\MiBaseDeDatos\MiBaseDeDatos.mdb;DefaultDir=" + App.Path + _
"\MiBaseDeDatos;DriverId=281;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"






'esta parte el cdigo Muestra Nuestro DataReport
DataReport1.Show
Exit Sub

'Esta Parte del Codigo es Para x Si Hay un Error Al Abrir el DataReport
ErrorOpen:

Unload DataReport1
DataReport1.Show



End Sub

Private Sub buscar_Click()
Dim m As Long

m = InputBox("Introduce la Matrícula que Buscas")

Data1.Recordset.FindFirst "matricula=" & m
If Data1.Recordset.NoMatch Then
MsgBox "La Matrícula Número: " & m & " No está en la Base de Datos", vbExclamation, "Búsquedas de Matrícula"
End If

End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
MsgBox "No Se Encontro la Carpeta:" _
+ "MiBaseDeDatos que es la que Contiene la Base de Datos", vbCritical, "Errorla conectar la base de Datos!!!!"
End
End Sub

Private Sub eliminar_Click()
If MsgBox("¿Quieres Eliminar la Matrícula Número: " & Text1 & "?", vbYesNo, "Eliminar Registro") = 6 Then
Data1.Recordset.Delete
Data1.Refresh
Text1.SetFocus
MsgBox "Se Eliminó la Matrícula", vbCritical, "Aviso Importante"
Else
MsgBox "No se Eliminó la Matrícula Número: " & Text1, vbExclamation, "Aviso Importante"
End If
End Sub

Private Sub Form_Load()
Label5.Caption = App.Path + "\" + App.EXEName
'Esta parte del codigo es para que el data1 _
se conecte ala base de datos ya creada segun el path dinde se encuentre

Data1.DatabaseName = App.Path + "\MiBaseDeDatos\MiBaseDeDatos.mdb"



End Sub

Private Sub guardar_Click()
Data1.UpdateRecord
Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
End Sub

Private Sub nuevo_Click()
Data1.Recordset.AddNew
End Sub
