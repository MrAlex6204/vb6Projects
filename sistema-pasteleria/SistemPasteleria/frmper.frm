VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmper 
   BackColor       =   &H8000000E&
   Caption         =   "Personal"
   ClientHeight    =   4095
   ClientLeft      =   3240
   ClientTop       =   930
   ClientWidth     =   6420
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6420
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   4440
      Picture         =   "frmper.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   3840
      Picture         =   "frmper.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      Height          =   300
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A"
      Height          =   300
      Left            =   2640
      TabIndex        =   17
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   2040
      Picture         =   "frmper.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   1440
      Picture         =   "frmper.frx":0B96
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   580
   End
   Begin MSDataGridLib.DataGrid dgper 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
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
            Format          =   ""
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
   Begin MSDataListLib.DataCombo Datacbodis 
      Height          =   315
      Left            =   960
      TabIndex        =   13
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo Datacbocar 
      Height          =   315
      Left            =   3960
      TabIndex        =   12
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtfec 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Width           =   2100
   End
   Begin VB.TextBox txttel 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3960
      TabIndex        =   10
      Top             =   1680
      Width           =   2100
   End
   Begin VB.TextBox txtdir 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
      Width           =   2100
   End
   Begin VB.TextBox txtnom 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   960
      TabIndex        =   8
      Top             =   1320
      Width           =   2100
   End
   Begin VB.TextBox txtdniper 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   2100
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ingreso"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distrito"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   285
   End
End
Attribute VB_Name = "frmper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim per As String

Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR AL PERSONAL NUEVO", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
copiarcampos
rsper.Update
Else
rsper.Cancel
End If
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rsper.Bookmark
b = InputBox("INGRESE NOMBRE DE LA PERSONA A BUSCAR", "SISTEMA DE SEGURIDAD")
rsper.MoveFirst
rsper.Find "nom_per='" + Trim(b) + "'"
If rsper.EOF Then
MsgBox "NOMBRE DE LA PERSONA NO EXISTE", vbCritical, "SISTEMA DE SEGURIDAD"
rsper.Bookmark = srt1
End If
llenarcampos
End Sub


Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("DESEA ELIMINAR A LA PERSONA", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rsper.Delete
rsper.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command4_Click()
txtdniper.SetFocus
End Sub

Private Sub Command5_Click()
limpiarcampos
rsper.AddNew
habilitacontrol
per = txtdniper.Text & (rsper.RecordCount)
per = Right(per, 8)



End Sub

Private Sub Command6_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA ", vbYesNo, "SISTEMA DE SEGURIDAD")
If X = vbYes Then
Unload Me
End If
frmper.Hide
End Sub

Private Sub Datacbocar_Click(Area As Integer)
rscar.MoveFirst
nom_car = Datacbocar.Text
rscar.Find "nom_car='" + Trim(Datacbocar.Text) + "'"
If rscar.EOF Then
Else
End If
End Sub

Private Sub Datacbodis_Click(Area As Integer)

rsdis.MoveFirst
nom_dis = Datacbodis.Text
rsdis.Find "nom_dis='" + Trim(Datacbodis.Text) + "'"
If rsdis.EOF Then
Else
End If
End Sub

Private Sub Form_Load()

activaper
llenarcampos


activadis
Datacbodis.ListField = "nom_dis"
Set Datacbodis.RowSource = rsdis

deshabilitacontrol

per = txtdniper.Text & (rsper.RecordCount)
per = Right(per, 8)

Set rsper = New ADODB.Recordset
rsper.CursorLocation = adUseClient
sqlper = "select *from personal"
rsper.Open sqlper, cn, adOpenStatic, adLockOptimistic
Set dgper.DataSource = rsper
dgper.Refresh
dgper.Columns(0).Visible = False
dgper.Columns(1).Visible = False
dgper.Columns(2).Caption = "Cargo"
dgper.Columns(3).Caption = "Nombre"
dgper.Columns(4).Visible = False
dgper.Columns(5).Visible = False
dgper.Columns(6).Caption = "Telefono"
dgper.Columns(7).Visible = False

dgper.Columns(2).Width = 0.2 * dgper.Width
dgper.Columns(3).Width = 0.47 * dgper.Width
dgper.Columns(6).Width = 0.22 * dgper.Width
activacar
Datacbocar.ListField = "nom_car"
Set Datacbocar.RowSource = rscar



End Sub


Public Sub llenarcampos()
If rsper.BOF Then Exit Sub
If rsper.EOF Then Exit Sub
txtdniper.Text = rsper.Fields("DNI_per")
Datacbocar.Text = rsper.Fields("nom_car")
txtnom.Text = rsper.Fields("nom_per")
txtdir.Text = rsper.Fields("dir_per")
Datacbodis.Text = rsper.Fields("cod_dis")
txttel.Text = rsper.Fields("tel_per")
txtfec.Text = rsper.Fields("fec_ingper")
End Sub
Public Sub limpiarcampos()
txtdniper.Text = ""
Datacbocar.Text = ""
txtnom.Text = ""
txtdir.Text = ""
Datacbodis.Text = ""
txttel.Text = ""
txtfec.Text = ""
End Sub

Public Sub copiarcampos()
 rsper.Fields("DNI_per") = txtdniper.Text
 rsper.Fields("nom_car") = Datacbocar.Text
 rsper.Fields("nom_per") = txtnom.Text
 rsper.Fields("dir_per") = txtdir.Text
 rsper.Fields("cod_dis") = Datacbodis.Text
 rsper.Fields("tel_per") = txttel.Text
 rsper.Fields("fec_ingper") = txtfec.Text

End Sub

Public Sub habilitacontrol()
txtdniper.Enabled = True
Datacbocar.Enabled = True
txtnom.Enabled = True
txtdir.Enabled = True
Datacbodis.Enabled = True
txttel.Enabled = True
txtfec.Enabled = True
End Sub

Public Sub deshabilitacontrol()

txtdniper.Enabled = False
Datacbocar.Enabled = False
txtnom.Enabled = False
txtdir.Enabled = False
Datacbodis.Enabled = False
txttel.Enabled = False
txtfec.Enabled = False

End Sub

