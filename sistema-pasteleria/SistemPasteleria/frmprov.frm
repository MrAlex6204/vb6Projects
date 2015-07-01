VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmprov 
   BackColor       =   &H8000000E&
   Caption         =   "Proveedores"
   ClientHeight    =   4185
   ClientLeft      =   3240
   ClientTop       =   690
   ClientWidth     =   6675
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6675
   Begin MSDataGridLib.DataGrid dgprov 
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
            Type            =   1
            Format          =   "dd. MMMM""ta ""yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
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
            LCID            =   2058
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
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   4320
      Picture         =   "frmprov.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   3720
      Picture         =   "frmprov.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      Height          =   300
      Left            =   3120
      TabIndex        =   22
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A"
      Height          =   300
      Left            =   2520
      TabIndex        =   21
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   1920
      Picture         =   "frmprov.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   1320
      Picture         =   "frmprov.frx":0B96
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   580
   End
   Begin MSDataListLib.DataCombo Datacbodis 
      Height          =   315
      Left            =   4320
      TabIndex        =   17
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtfecing 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   16
      Top             =   2280
      Width           =   2220
   End
   Begin VB.TextBox txtpagprov 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4320
      TabIndex        =   15
      Top             =   1920
      Width           =   2220
   End
   Begin VB.TextBox txtemaprov 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   1920
      Width           =   2220
   End
   Begin VB.TextBox txtrucprov 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4320
      TabIndex        =   13
      Top             =   1560
      Width           =   2220
   End
   Begin VB.TextBox txttel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   12
      Top             =   1560
      Width           =   2220
   End
   Begin VB.TextBox txtdir 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   2220
   End
   Begin VB.TextBox txtnomprov 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   2220
   End
   Begin VB.TextBox txtdniprov 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   9
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ingreso"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Web"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distritos"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   285
   End
End
Attribute VB_Name = "frmprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR AL PROVEEDOR", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
'--------------------
copiarcampos
rsprov.Update
'-----------------------
Else
rspro.Cancel
End If
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rsprov.Bookmark
b = InputBox("INGRESE NOMBRE DE PROVEEDOR A BUSCAR", "SISTEMA DE SEGURIDAD")
rsprov.Find "nom_prov='" + Trim(b) + "' "
If rspro.EOF Then
MsgBox "NOMBRE DE PROVEEDOR A BUSCAR NO EXIXTE", vbCritical, "SISTEMA DE SEGURIDAD"
rsprov.Bookmark = str1
End If
llenarcamposprov
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("DESEA ANULAR EL PROVEEDOR", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rsprov.Delete
rsprov.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command4_Click()

txtdniprov.SetFocus
End Sub

Private Sub Command5_Click()
limpiarcampos
'-----------------------------
rsprov.AddNew
'---------------------------
txtdniprov.SetFocus
End Sub

Private Sub Command6_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DE LA TABLA PROVEEDORES", vbYesNo, "SISTEMA DE SEGURIDAD")
If X = vbYes Then
Unload Me
frmprov.Hide
End If

End Sub

Private Sub datacbodis_Click(Area As Integer)
rsdis.MoveFirst
nom_dis = datacbodis.Text
rsdis.Find "nom_dis='" + Trim(datacbodis.Text) + "'"
If rsdis.EOF Then
Else
End If
End Sub

Private Sub Form_Load()
activaprov
llenarcamposprov

activadis
datacbodis.ListField = "nom_dis"
Set datacbodis.RowSource = rsdis



Set rsprov = New ADODB.Recordset
rsprov.CursorLocation = adUseClient
sqlprov = "select *from proveedores"
rsprov.Open sqlprov, cn, adOpenStatic, adLockOptimistic

Set dgprov.DataSource = rsprov
dgprov.Refresh
dgprov.Columns(0).Visible = False
dgprov.Columns(1).Caption = "Proveedor"
dgprov.Columns(2).Caption = "Direccion"
dgprov.Columns(3).Visible = False
dgprov.Columns(4).Visible = False
dgprov.Columns(5).Visible = False
dgprov.Columns(6).Visible = False
dgprov.Columns(7).Caption = "Web Sities"
dgprov.Columns(8).Visible = False
dgprov.Columns(1).Width = 0.32 * dgprov.Width
dgprov.Columns(2).Width = 0.3 * dgprov.Width
dgprov.Columns(7).Width = 0.32 * dgprov.Width

End Sub

Public Sub llenarcamposprov()
If rsprov.BOF Then Exit Sub
If rsprov.EOF Then Exit Sub

txtdniprov.Text = rsprov.Fields("DNI_prov")
txtnomprov.Text = rsprov.Fields("nom_prov")
txtdir.Text = rsprov.Fields("dir_prov")
datacbodis.Text = rsprov.Fields("cod_dis")
txttel.Text = rsprov.Fields("tel_prov")
txtrucprov.Text = rsprov.Fields("RUC_prov")
txtemaprov.Text = rsprov.Fields("ema_prov")
txtpagprov.Text = rsprov.Fields("pag_web")
txtfecing.Text = rsprov.Fields("fec_ing")
End Sub
'-----------------------------------------------------
Public Sub copiarcampos()

rsprov.Fields("DNI_prov") = txtdniprov.Text
rsprov.Fields("nom_prov") = txtnomprov.Text
rsprov.Fields("dir_prov") = txtdir.Text
rsprov.Fields("cod_dis") = datacbodis.Text
rsprov.Fields("tel_prov") = txttel.Text
rsprov.Fields("RUC_prov") = txtrucprov.Text
rsprov.Fields("ema_prov") = txtemaprov.Text
rsprov.Fields("pag_web") = txtpagprov.Text
rsprov.Fields("fec_ing") = txtfecing.Text

'--------------------------------------------------
End Sub

Public Sub limpiarcampos()

txtdniprov.Text = ""
txtnomprov.Text = ""
txtdir.Text = ""
datacbodis.Text = ""
txttel.Text = ""
txtrucprov.Text = ""
txtemaprov.Text = ""
txtpagprov.Text = ""
txtfecing.Text = ""
End Sub

