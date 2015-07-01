VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcmarcas 
   BackColor       =   &H80000009&
   Caption         =   "Marcas"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgmar 
      Height          =   2175
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSDataListLib.DataList dbllista 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3889
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Marcas"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "frmcmarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nom_mar As String

Private Sub Command1_Click()
Unload Me
frmcmarcas.Hide
End Sub

Private Sub Command2_Click()
dtamarcas.Show
End Sub

Private Sub dbllista_Click()

       ' Llena el DATACOMBO (dtcproducto) con los nombre de los articulos según su categoría
        
        Set rspro = New ADODB.Recordset
        
        rspro.CursorLocation = adUseClient
        rspro.Open "select *from productos where cod_mar='" + Trim(dbllista.BoundText) + "'", cn, adOpenStatic, adLockOptimistic
        
       Set dgmar.DataSource = rspro
dgmar.Columns(0).Visible = False
dgmar.Columns(1).Caption = "Producto"
dgmar.Columns(2).Visible = False
dgmar.Columns(3).Visible = False

dgmar.Columns(4).Visible = False
dgmar.Columns(5).Visible = False
dgmar.Columns(6).Visible = False
dgmar.Columns(7).Caption = "Precio"
dgmar.Columns(8).Caption = "Cantidad"
dgmar.Columns(9).Visible = False
dgmar.Columns(10).Visible = False
dgmar.Columns(11).Caption = "Empresa"

dgmar.Columns(1).Width = 0.35 * dgmar.Width
dgmar.Columns(7).Width = 0.1 * dgmar.Width
dgmar.Columns(8).Width = 0.15 * dgmar.Width
End Sub

Private Sub Form_Load()
activamar


Set dbllista.RowSource = rsmar
dbllista.ListField = "nom_mar"
dbllista.BoundColumn = "cod_mar"

End Sub
