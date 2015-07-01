VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmveh 
   BackColor       =   &H80000009&
   Caption         =   "Vehiculos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dgveh 
      Height          =   1575
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2778
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
   Begin MSDataListLib.DataList dblveh 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2858
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Vehiculo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmveh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo As String

Private Sub Command1_Click()
Unload Me
frmveh.Hide
End Sub

Private Sub dblveh_Click()

' Llena el DATACOMBO (dtcproducto) con los nombre de los articulos según su categoría
        
        Set rsve = New ADODB.Recordset
        
        rsve.CursorLocation = adUseClient
        rsve.Open "select *from transportistas where num_veh='" + Trim(dblveh.BoundText) + "'", cn, adOpenStatic, adLockOptimistic
        
              Set dgveh.DataSource = rsve
dgveh.Columns(0).Visible = False

dgveh.Columns(1).Caption = "Transportista"
dgveh.Columns(2).Caption = "Direccion"
dgveh.Columns(3).Visible = False
dgveh.Columns(4).Caption = "Placa"

dgveh.Columns(1).Width = 0.35 * dgveh.Width
dgveh.Columns(2).Width = 0.3 * dgveh.Width


End Sub

Private Sub Form_Load()
activave


Set dblveh.RowSource = rsve
dblveh.ListField = "tipo"
dblveh.BoundColumn = "num_veh"

End Sub
