VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpro 
   BackColor       =   &H8000000E&
   Caption         =   "Productos"
   ClientHeight    =   5265
   ClientLeft      =   3165
   ClientTop       =   450
   ClientWidth     =   8175
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8175
   Begin MSDataListLib.DataCombo Datacbomed 
      Height          =   315
      Left            =   1440
      TabIndex        =   32
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtdescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   5280
      Picture         =   "frmpro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   4680
      Picture         =   "frmpro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      Height          =   300
      Left            =   4080
      TabIndex        =   28
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A"
      Height          =   300
      Left            =   3480
      TabIndex        =   27
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   2880
      Picture         =   "frmpro.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   2280
      Picture         =   "frmpro.frx":0B96
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   120
      Width           =   580
   End
   Begin MSDataGridLib.DataGrid dgpro 
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
            Format          =   "dd-MMMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
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
   Begin MSDataListLib.DataCombo datacbomar 
      Height          =   315
      Left            =   5280
      TabIndex        =   23
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo datacbolin 
      Height          =   315
      Left            =   1440
      TabIndex        =   22
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtemp 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   21
      Top             =   3480
      Width           =   2700
   End
   Begin VB.TextBox txtnomp 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5280
      TabIndex        =   20
      Top             =   3120
      Width           =   2700
   End
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   19
      Top             =   3120
      Width           =   2700
   End
   Begin VB.TextBox txtstock 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5280
      TabIndex        =   18
      Top             =   2760
      Width           =   2700
   End
   Begin VB.TextBox txtstockm 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   17
      Top             =   2760
      Width           =   2700
   End
   Begin VB.TextBox txtprev 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5280
      TabIndex        =   16
      Top             =   2400
      Width           =   2700
   End
   Begin VB.TextBox txtprec 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   2700
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   14
      Top             =   1320
      Width           =   2700
   End
   Begin VB.TextBox txtcod 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   960
      Width           =   2700
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI Proveedor"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4200
      TabIndex        =   9
      Top             =   2760
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Venta"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Costo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linea"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medida"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caracteristica"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR EL PRODUCTO", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
copiarcampos
rspro.Update
Else
rspro.Cancel
End If
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rspro.Bookmark
b = InputBox("INGRESE NOMBRE DE LA PRODUCTO A BUSCAR", "SISTEMA DE SEGURIDAD")
rspro.MoveFirst
rspro.Find "des_pro='" + Trim(b) + "'"
If rspro.EOF Then
MsgBox "NOMBRE DEL PRODUCTO NO EXISTE", vbCritical, "SISTEMA DE SEGURIDAD"
rspro.Bookmark = srt1
End If
llenarcampospro
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("DESEA ANULAR EL PRODUCTO", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rspro.Delete
rspro.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command4_Click()
txtcod.SetFocus

End Sub

Private Sub Command5_Click()
limpiarcampos
rspro.AddNew
txtcod.SetFocus
End Sub

Private Sub Command6_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA", vbYesNo, "SISTEMA DE SEGURIDAD")
If X = vbYes Then
Unload Me
End If
frmpro.Hide
End Sub

Private Sub datacbolin_Click(Area As Integer)
rslin.MoveFirst
nom_lin = datacbolin.Text
rslin.Find "nom_lin='" + Trim(datacbolin.Text) + "'"
If rslin.EOF Then
Else
End If
End Sub

Private Sub datacbomar_Click(Area As Integer)
rsmar.MoveFirst
nom_mar = datacbomar.Text
rsmar.Find "nom_mar='" + Trim(datacbomar.Text) + "'"
If rsmar.EOF Then
Else
End If
End Sub

Private Sub Datacbomed_Click(Area As Integer)
rsmed.MoveFirst
nom_med = Datacbomed.Text
rsmed.Find "nom_med='" + Trim(Datacbomed.Text) + "'"
If rsmed.EOF Then
Else
End If
End Sub

Private Sub Form_Load()
activapro
llenarcampospro

activamar
datacbomar.ListField = "nom_mar"
Set datacbomar.RowSource = rsmar

activalin
datacbolin.ListField = "nom_lin"
Set datacbolin.RowSource = rslin

activamed
Datacbomed.ListField = "nom_med"
Set Datacbomed.RowSource = rsmed



Set rspro = New ADODB.Recordset
rspro.CursorLocation = adUseClient
sqlpro = "select *from productos"
rspro.Open sqlpro, cn, adOpenStatic, adLockOptimistic

Set dgpro.DataSource = rspro
dgpro.Refresh
dgpro.Columns(0).Visible = False
dgpro.Columns(1).Caption = "Descripcion"
dgpro.Columns(2).Visible = False
dgpro.Columns(3).Caption = "Medida"
dgpro.Columns(4).Caption = "Lineas"
dgpro.Columns(5).Visible = False
dgpro.Columns(6).Visible = False
dgpro.Columns(7).Visible = False
dgpro.Columns(8).Visible = False
dgpro.Columns(9).Visible = False
dgpro.Columns(10).Visible = False
dgpro.Columns(11).Caption = "Proveedor"
dgpro.Columns(12).Caption = "Empresa"
dgpro.Columns(1).Width = 0.22 * dgpro.Width
dgpro.Columns(3).Width = 0.12 * dgpro.Width
dgpro.Columns(4).Width = 0.13 * dgpro.Width
dgpro.Columns(11).Width = 0.25 * dgpro.Width
dgpro.Columns(12).Width = 0.2 * dgpro.Width

End Sub


Public Sub llenarcampospro()
If rspro.BOF Then Exit Sub
If rspro.EOF Then Exit Sub

txtcod.Text = rspro.Fields("cod_pro")
txtdes.Text = rspro.Fields("des_pro")
Txtdescripcion.Text = rspro.Fields("caracteristica")
Datacbomed = rspro.Fields("cod_uni")
datacbolin.Text = rspro.Fields("cod_lin")
datacbomar.Text = rspro.Fields("cod_mar")
txtprec.Text = rspro.Fields("pre_cos")
txtprev.Text = rspro.Fields("pre_ven")
txtstockm.Text = rspro.Fields("stock_min")
txtstock.Text = rspro.Fields("stock")
txtDNI.Text = rspro.Fields("DNI_prov")
txtnomp.Text = rspro.Fields("nom_prov")
txtemp.Text = rspro.Fields("empresa")
End Sub


Public Sub copiarcampos()
rspro.Fields("cod_pro") = txtcod.Text
rspro.Fields("des_pro") = txtdes.Text
rspro.Fields("caracteristica") = Txtdescripcion.Text
rspro.Fields("cod_uni") = datacbomrd
rspro.Fields("cod_lin") = datacbolin.Text
rspro.Fields("cod_mar") = datacbomar.Text
rspro.Fields("pre_cos") = txtprec.Text
rspro.Fields("pre_ven") = txtprev.Text
rspro.Fields("stock_min") = txtstockm.Text
rspro.Fields("stock") = txtstock.Text
rspro.Fields("DNI_prov") = txtDNI.Text
rspro.Fields("nom_prov") = txtnomp.Text
rspro.Fields("empresa") = txtemp.Text
End Sub

Public Sub limpiarcampos()
txtcod.Text = ""
txtdes.Text = ""
Txtdescripcion.Text = ""
Datacbomed.Text = ""
datacbolin.Text = ""
datacbomar.Text = ""
txtprec.Text = ""
txtprev.Text = ""
txtstockm.Text = ""
txtstock.Text = ""
txtDNI.Text = ""
txtnomp.Text = ""
txtemp.Text = ""
End Sub

