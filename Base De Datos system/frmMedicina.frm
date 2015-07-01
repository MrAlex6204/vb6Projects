VERSION 5.00
Begin VB.Form frmMedicina 
   BackColor       =   &H00000000&
   Caption         =   "MEDICINA"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10755
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmMedicina.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "CONSULTA"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\usurio\Escritorio\PRACTIVANTE 1\BASEDATOS\BaseDeDatos\Medicina.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Datos"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Matricula"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Apellidos"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "Turno"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   7200
      Picture         =   "frmMedicina.frx":0E72
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2700
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Medicina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "APELLIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   930
   End
   Begin VB.Menu MovMed 
      Caption         =   "&Movimientos"
      WindowList      =   -1  'True
      Begin VB.Menu AltaMed 
         Caption         =   "&Alta"
      End
      Begin VB.Menu EditMed 
         Caption         =   "Editar"
      End
      Begin VB.Menu BuscarMed 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu EliminarMed 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu GuardarMed 
         Caption         =   "&Guardar"
      End
   End
   Begin VB.Menu SalirMed 
      Caption         =   "&Salir"
      Begin VB.Menu CerrarMed 
         Caption         =   "&Cerrar Medicina"
      End
   End
End
Attribute VB_Name = "frmMedicina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AltaMed_Click()

Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True

Command1.Visible = True
Command2.Visible = True

Data1.Recordset.AddNew
End Sub

Private Sub BuscarMed_Click()
On Error GoTo salir
Dim m As Long

m = InputBox("Introduce la Matrícula que Buscas")

Data1.Recordset.FindFirst "Matricula='" & m & "'"
If Data1.Recordset.NoMatch Then
MsgBox "La Matrícula Número: " & m & " No está en la Base de Datos", vbExclamation, "Búsquedas de Matrícula"
End If
Exit Sub
salir:
End Sub

Private Sub CerrarMed_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Text1 = Empty Then
MsgBox "Introdusca la Matricula", vbExclamation, "Aviso Importante"
Text1.SetFocus
Exit Sub
End If
If Text2 = Empty Then
MsgBox "Introdusca el Nombre", vbExclamation, "Aviso Importante"
Text2.SetFocus
Exit Sub
End If
If Text3 = Empty Then
MsgBox "Introdusca los Apellidos", vbExclamation, "Aviso Importante"
Text3.SetFocus
Exit Sub
End If
If Text4 = Empty Then
MsgBox "Introdusca el Turno", vbExclamation, "Aviso Importante"
Text4.SetFocus
Exit Sub
End If

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False

Command1.Visible = False
Command2.Visible = False


Data1.UpdateRecord
Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"

Unload Me
frmMedicina.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command3.Visible = False
End Sub

Private Sub EditMed_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command3.Visible = True
Command3.Top = 4200
Command3.Left = 4200
End Sub

Private Sub EliminarMed_Click()
If MsgBox("¿Quieres Eliminar la Matrícula Número: " & Text1 & "?", vbYesNo, "Eliminar Registro") = 6 Then
Data1.Recordset.Delete
Data1.Refresh

MsgBox "Se Eliminó la Matrícula", vbCritical, "Aviso Importante"
Else
MsgBox "No se Eliminó la Matrícula Número: " & Text1, vbExclamation, "Aviso Importante"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Data1.DatabaseName = App.Path + "\Medicina.mdb"
End Sub

Private Sub GuardarMed_Click()
Data1.UpdateRecord
Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
End Sub

Private Sub SalirM_Click()
End
End Sub
