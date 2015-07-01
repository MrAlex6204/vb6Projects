VERSION 5.00
Begin VB.Form Dialog1 
   BackColor       =   &H00000000&
   Caption         =   "Total de Ventas"
   ClientHeight    =   4440
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11085
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Reporte de Ventas del Cajero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   3855
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   8295
      Begin VB.CommandButton Command7 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Articulos Vendidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   525
      End
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
   Public RS As New ADODB.Recordset
   Public RVentArt As New ADODB.Recordset
   
Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset  'Objeto para manejar la Tabla
Dim SQL As String
 Dim strConnect As String
 Dim strPath As String
 
Private Sub Command7_Click()
Unload Me
End Sub

Sub AbrirBase()
RS.Open "select * from  VentReg", strConnect, adOpenKeyset, adLockOptimistic
RVentArt.Open "select * from VentArt", strConnect, adOpenKeyset, adLockOptimistic
End Sub
Sub CerrarBase()
RS.Close
RVentArt.Close
End Sub

Private Sub Command8_Click()
SQL = "DELETE FROM VentReg"
BDD.Execute SQL

End Sub

Private Sub Form_Load()
On Error Resume Next
strPath = App.Path & "\Tienda.mdb"

strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
     "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      
      
Set BDD = OpenDatabase(App.Path & "\Tienda.mdb")


AbrirBase

SQL = "SELECT COUNT(cajero) FROM  VentReg"   'para saber la cantidad de registros (incluye los nulos)
Set TBL = BDD.OpenRecordset(SQL)
Text7 = TBL("expr1000") 'expr1000 es el name del item de TBL q almacena el resultado del operador de agrupamiento



SQL = "SELECT SUM(Total) FROM  VentReg"    'sumatoria de los precios
Set TBL = BDD.OpenRecordset(SQL)

Text9 = TBL("expr1000")

CerrarBase
End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub OKButton_Click()

End Sub
