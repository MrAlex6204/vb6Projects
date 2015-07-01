VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcdistrito 
   BackColor       =   &H8000000E&
   Caption         =   "Distritos"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgdis 
      Height          =   2175
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
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
   Begin MSDataListLib.DataList dbldis 
      Height          =   2205
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3889
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Distritos"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmcdistrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dbldis_Click()
' Llena el DATACOMBO (dtcproducto) con los nombre de los articulos según su categoría
        
        Set rscli = New ADODB.Recordset
        
        rscli.CursorLocation = adUseClient
        rscli.Open "select *from clientes where cod_dis='" + Trim(dbldis.BoundText) + "'", cn, adOpenStatic, adLockOptimistic
        
       Set dgdis.DataSource = rscli
dgdis.Columns(0).Visible = False
dgdis.Columns(1).Caption = "Cliente"
dgdis.Columns(2).Caption = "Direccion"
dgdis.Columns(3).Visible = False
dgdis.Columns(4).Caption = "Telefono"
dgdis.Columns(5).Visible = False
dgdis.Columns(6).Visible = False
dgdis.Columns(7).Visible = False
dgdis.Columns(1).Width = 0.4 * dgdis.Width
dgdis.Columns(4).Width = 0.15 * dgdis.Width


End Sub

Private Sub Form_Load()
activadis
Set dbldis.RowSource = rsdis
dbldis.ListField = "nom_dis"
dbldis.BoundColumn = "cod_dis"

End Sub
