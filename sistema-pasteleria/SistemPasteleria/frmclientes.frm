VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmclientes 
   BackColor       =   &H8000000E&
   Caption         =   "Clientes"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optbuscar 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optbuscar 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgcli 
      Height          =   1935
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.TextBox txtbuscar 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por"
      Height          =   195
      Left            =   6960
      TabIndex        =   5
      Top             =   840
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   8280
      TabIndex        =   2
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dni"
      Height          =   195
      Left            =   7200
      TabIndex        =   1
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblbuscar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   45
   End
End
Attribute VB_Name = "frmclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim campo As String

Private Sub Command1_Click()
Unload Me
frmclientes.Hide
End Sub

Private Sub Form_Load()
Set rscli = New ADODB.Recordset
rscli.CursorLocation = adUseClient
sqlcli = "select DNI_cli,nom_cli,dir_cli,fec_ing from clientes"
rscli.Open sqlcli, cn, adOpenStatic, adLockOptimistic
Set dgcli.DataSource = rscli
dgcli.Refresh




campo = "DNI_cli"
End Sub



Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
lblbuscar.Caption = "Buscar por DNI"
campo = "DNI_cli"

Case 1
lblbuscar.Caption = "Buscar por Nombre del Cliente"
campo = "nom_cli"
End Select
txtbuscar.Enabled = True
txtbuscar.SetFocus
End Sub


Private Sub optbuscar_Click(Index As Integer)
Select Case Index
Case 0
lblbuscar.Caption = "Buscar por DNI"
campo = "DNI_cli"

Case 1
lblbuscar.Caption = "Buscar por Nombre del Cliente"
campo = "nom_cli"
End Select
txtbuscar.Enabled = True
txtbuscar.SetFocus

End Sub

Private Sub txtbuscar_Change()

Dim sqlbus As String
rscli.Close
sqlbus = "select *from clientes" & " where " + Trim(campo) + " like '" + Trim(txtbuscar.Text) + "%'" & "order by nom_cli "
rscli.Open sqlbus, cn, adOpenStatic, adLockOptimistic
Set dgcli.DataSource = rscli
dgcli.Refresh





End Sub
