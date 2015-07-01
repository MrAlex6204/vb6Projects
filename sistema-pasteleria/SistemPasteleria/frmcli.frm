VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcli 
   BackColor       =   &H8000000E&
   Caption         =   "Clientes"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgcli 
      Height          =   1095
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1931
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      RowDividerStyle =   1
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
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   4200
      Picture         =   "frmcli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   4800
      Picture         =   "frmcli.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A"
      Height          =   300
      Left            =   3000
      TabIndex        =   19
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E"
      Height          =   300
      Left            =   3600
      TabIndex        =   18
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   2400
      Picture         =   "frmcli.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   1800
      Picture         =   "frmcli.frx":0B96
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   580
   End
   Begin MSDataListLib.DataCombo Datacbodis 
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtfec 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   14
      Top             =   1920
      Width           =   2220
   End
   Begin VB.TextBox txtema 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4800
      TabIndex        =   13
      Top             =   1920
      Width           =   2220
   End
   Begin VB.TextBox txtRUCcli 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   2220
   End
   Begin VB.TextBox txttelcli 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   2220
   End
   Begin VB.TextBox txtdircli 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   2220
   End
   Begin VB.TextBox txtnomcli 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4800
      TabIndex        =   9
      Top             =   840
      Width           =   2220
   End
   Begin VB.TextBox txtDNIcli 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ingreso"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distrito"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   285
   End
End
Attribute VB_Name = "frmcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dni As String
Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR AL CLIENTE", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
grabarcampos
rscli.Update
Else
rscli.Cancel
End If
deshabilitacampos
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rscli.Bookmark
b = InputBox("INGRESE NOMBRE DEL CLIENTE A BUSCAR", "SISTEMA DE SEGURIDAD")
rscli.MoveFirst
rscli.Find "dni_cli='" + Trim(b) + "'"
If rscli.EOF Then
MsgBox "NOMBRE DEL CLIENTE NO EXISTE", vbCritical, "SISTEMA DE SEGURIDAD"
rscli.Bookmark = str1
End If
llenarcamposcli
End Sub

Private Sub Command3_Click()
txtDNIcli.SetFocus

End Sub

Private Sub Command4_Click()
Dim a As Integer
a = MsgBox("DESEA ELIMINAR AL CLIENTE", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rscli.Delete
rscli.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command5_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA", vbYesNo, "SISTEMA DE SEGURIDAD")
If vbYes Then
Unload Me
frmcli.Hide
End If

End Sub

Private Sub Command6_Click()
limpiarcampos
rscli.AddNew

dni = txtDNIcli.Text & (rscli.RecordCount)
dni = Right(DNI_cli, 8)
habilitacampos

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

activacli
llenarcamposcli
deshabilitacampos


dni = txtDNIcli.Text & (rscli.RecordCount)
dni = Right(DNI_cli, 8)

activadis

Datacbodis.ListField = "nom_dis"
Set Datacbodis.RowSource = rsdis


Set rscli = New ADODB.Recordset
rscli.CursorLocation = adUseClient
sqlcli = "select *from clientes"
rscli.Open sqlcli, cn, adOpenStatic, adLockOptimistic
Set dgcli.DataSource = rscli
dgcli.Refresh

dgcli.Columns(0).Visible = False
dgcli.Columns(1).Caption = "Cliente"
dgcli.Columns(2).Visible = False
dgcli.Columns(3).Visible = False
dgcli.Columns(4).Caption = "Telefono"
dgcli.Columns(5).Visible = False
dgcli.Columns(6).Visible = False
dgcli.Columns(7).Caption = "Ingreso"
dgcli.Columns(1).Width = 0.4 * dgcli.Width
dgcli.Columns(4).Width = 0.18 * dgcli.Width
dgcli.Columns(7).Width = 0.3 * dgcli.Width

End Sub


Public Sub llenarcamposcli()
If rscli.BOF Then Exit Sub
If rscli.EOF Then Exit Sub

txtDNIcli.Text = rscli.Fields("DNI_cli")
txtnomcli.Text = rscli.Fields("nom_cli")
txtdircli.Text = rscli.Fields("dir_cli")
Datacbodis.Text = rscli.Fields("cod_dis")
txttelcli.Text = rscli.Fields("tel_cli")

txtRUCcli.Text = rscli.Fields("RUC_cli")
txtema.Text = rscli.Fields("ema_cli")
txtfec.Text = rscli.Fields("fec_ing")
End Sub


Public Sub grabarcampos()
rscli.Fields("DNI_cli") = txtDNIcli.Text
rscli.Fields("nom_cli") = txtnomcli.Text
rscli.Fields("dir_cli") = txtdircli.Text
rscli.Fields("cod_dis") = Datacbodis.Text
rscli.Fields("tel_cli") = txttelcli.Text

rscli.Fields("RUC_cli") = txtRUCcli.Text
rscli.Fields("ema_cli") = txtema.Text
rscli.Fields("fec_ing") = txtfec.Text
End Sub

Public Sub limpiarcampos()
txtDNIcli.Text = ""
txtnomcli.Text = ""
txtdircli.Text = ""
Datacbodis.Text = ""
txttelcli.Text = ""
txtRUCcli.Text = ""
txtema.Text = ""
txtfec.Text = ""
End Sub
Public Sub habilitacampos()
txtDNIcli.Enabled = True
txtnomcli.Enabled = True
txtdircli.Enabled = True
Datacbodis.Enabled = True
txttelcli.Enabled = True
txtRUCcli.Enabled = True
txtema.Enabled = True
txtfec.Enabled = True

End Sub

Public Sub deshabilitacampos()
txtDNIcli.Enabled = False
txtnomcli.Enabled = False
txtdircli.Enabled = False
Datacbodis.Enabled = False
txttelcli.Enabled = False
txtRUCcli.Enabled = False
txtema.Enabled = False
txtfec.Enabled = False

End Sub

