VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmproveedores 
   BackColor       =   &H8000000E&
   Caption         =   "Proveedores"
   ClientHeight    =   4320
   ClientLeft      =   135
   ClientTop       =   4215
   ClientWidth     =   9015
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   9015
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgprov 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      AllowDelete     =   -1  'True
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
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5175
   End
   Begin VB.OptionButton optpro 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   4
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optpro 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "PANADERIA PASTELERIA ALISSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   4905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7440
      TabIndex        =   6
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblbuscar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "frmproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim campo As String

Private Sub Command1_Click()
Unload Me
frmproveedores.Hide
End Sub

Private Sub Form_Load()

Set rsprov = New ADODB.Recordset
rsprov.CursorLocation = adUseClient
sqlprov = "select DNI_prov,nom_prov,dir_prov,pag_web,fec_ing from proveedores"
rsprov.Open sqlprov, cn, adOpenStatic, adLockOptimistic
Set dgprov.DataSource = rsprov
dgprov.Refresh


campo = "nom_prov"

End Sub





Private Sub optpro_Click(Index As Integer)
Select Case Index
Case 0
lblbuscar.Caption = "Buscar por Proveedor"
campo = "nom_prov"

Case 1
lblbuscar.Caption = "Buscar por Numero de Ruc"
campo = "RUC_prov"
End Select
txtbuscar.Enabled = True
txtbuscar.SetFocus
End Sub


Private Sub txtbuscar_Change()
Dim sqlbus As String
rsprov.Close
sqlbus = "select *from consulta" & " where " + Trim(campo) + " like '" + Trim(txtbuscar.Text) + "%'" & "order by nom_prov "
rsprov.Open sqlbus, cn, adOpenStatic, adLockOptimistic
Set dgprov.DataSource = rsprov
dgprov.Refresh


End Sub
