VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmproducto 
   BackColor       =   &H8000000E&
   Caption         =   "Productos"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtbus 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin MSDataGridLib.DataGrid dgproducto 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3836
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
   Begin VB.Label lblbuscar 
      BackColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Empresa"
      Height          =   195
      Left            =   7080
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Producto"
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   1320
      Width           =   645
   End
End
Attribute VB_Name = "frmproducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim campo As String
Private Sub Command1_Click()
Unload Me
frmproducto.Hide
End Sub

Private Sub Form_Load()
Set rspro = New ADODB.Recordset
rspro.CursorLocation = adUseClient
sqlpro = "select  *from productos"
rspro.Open sqlpro, cn, adOpenStatic, adLockOptimistic
Set dgproducto.DataSource = rspro
dgproducto.Refresh
dgproducto.Columns(0).Visible = False
dgproducto.Columns(1).Caption = "Producto"
dgproducto.Columns(2).Visible = False
dgproducto.Columns(3).Visible = False
dgproducto.Columns(4).Visible = False
dgproducto.Columns(5).Visible = False
dgproducto.Columns(6).Visible = False
dgproducto.Columns(7).Caption = "Precio Costo"
dgproducto.Columns(8).Caption = "Stock"
dgproducto.Columns(9).Visible = False
dgproducto.Columns(10).Visible = False
dgproducto.Columns(11).Caption = "Empresa"



campo = "nom_prov"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
lblbuscar.Caption = "Buscar Producto"
campo = "des_pro"

Case 1
lblbuscar.Caption = "Buscar por Linea"
campo = "empresa"

End Select
txtbus.Enabled = True
txtbus.SetFocus
End Sub

Private Sub txtbus_Change()
Dim sqlbus As String
rspro.Close
sqlbus = "select *from productos" & " where " + Trim(campo) + " like '" + Trim(txtbus.Text) + "%'" & "order by des_pro "
rspro.Open sqlbus, cn, adOpenStatic, adLockOptimistic
Set dgproducto.DataSource = rspro
dgproducto.Refresh
End Sub
