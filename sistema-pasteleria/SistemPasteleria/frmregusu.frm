VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmregusu 
   BackColor       =   &H8000000E&
   Caption         =   "Registro de los Usuarios"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgvista 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4895
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de los Usuarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   3120
   End
End
Attribute VB_Name = "frmregusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
rsreg.Close
Set rsreg = New ADODB.Recordset
rsreg.CursorLocation = adUseClient
sqlreg = "select *from registros"
rsreg.Open sqlreg, cn, adOpenStatic, adLockOptimistic

Set dgvista.DataSource = rsreg

dgvista.Refresh
dgvista.Columns(0).Caption = "Nombre" '
dgvista.Columns(1).Caption = "Cargo"
dgvista.Columns(2).Caption = "Fecha Ingreso"
dgvista.Columns(3).Caption = "Hora"
dgvista.Columns(4).Caption = "Fecha Salida"
dgvista.Columns(0).Width = 0.3 * dgvista.Width
dgvista.Columns(1).Width = 0.2 * dgvista.Width
End Sub
