VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Base Datos Ejemplo.."
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "Clave"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   2  'Align Bottom
      Bindings        =   "Form1.frx":0000
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   3900
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0015
      OLEDBString     =   $"Form1.frx":00B0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Maestros"
      Caption         =   "Consulta de Datos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      Begin VB.TextBox Text1 
         DataField       =   "Clave"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         DataField       =   "Nombre"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         DataField       =   "Apellido"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apellido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   1260
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BaseDatos As New ADODB.Recordset

Private Sub Command1_Click()

Form2.Show
End Sub

Private Sub Form_Load()
Dim sItemData As String
   Dim strData As String
   Dim strOutData As String
   Dim ConexionDatos As String
   
        
   
   
   
    Dim strPath As String
     
    'path donde se encuentra la base de datos
    strPath = "E:\escuela\Visual Basic\BaseDatos\BaseDatos.mdb"
    
   ConexionDatos = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
   ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    
    
   ' strData es para buscar :Campo='Valor a Buscar'
    strData = "Clave = '1055'"
    
    BaseDatos.Open "select * from Maestros", ConexionDatos, adOpenKeyset, adLockOptimistic
    BaseDatos.Find strData
    'Para Saber si no se encuntra un registro
    If BaseDatos.EOF = True Then
    MsgBox "No Se Encuentra el Registro:" & strData & " En la tabla", vbExclamation, "No se Encontro!!"
    
    End If
    On Error GoTo salir:
    
    'Busca en el Campo Clave de La Tabla
    strOutData = BaseDatos.Fields("Clave")
    Text1.Text = strOutData
    
    strOutData = BaseDatos.Fields("Nombre")
    Text2.Text = strOutData
    
    strOutData = BaseDatos.Fields("Apellido")
    Text3.Text = strOutData
    
salir:
     

    
    
End Sub
