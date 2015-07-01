VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcar 
   BackColor       =   &H80000009&
   Caption         =   "Cargos"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   1560
      Picture         =   "frmcar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   580
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   2160
      Picture         =   "frmcar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E"
      Height          =   300
      Left            =   3360
      TabIndex        =   8
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A"
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   4560
      Picture         =   "frmcar.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   580
   End
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   3960
      Picture         =   "frmcar.frx":1156
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   580
   End
   Begin MSDataGridLib.DataGrid dgcar 
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1931
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
   Begin VB.TextBox txtcar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtcod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Cargo"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Codigo"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod As String

Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR AL CLIENTE", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
grabarcampos
rscar.Update
Else
rscar.Cancel
End If
End Sub

Public Sub grabarcampos()

rscar.Fields("cod_car") = txtcod.Text
rscar.Fields("nom_car") = txtcar.Text
End Sub

Public Sub limpiarcampos()
txtcod.Text = ""
txtcar.Text = ""
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rscar.Bookmark
b = InputBox("INGRESE NOMBRE DEL CLIENTE A BUSCAR", "SISTEMA DE SEGURIDAD")
rscar.MoveFirst
rscar.Find "cod_car='" + Trim(b) + "'"
If rscar.EOF Then
MsgBox "NOMBRE DEL CLIENTE NO EXISTE", vbCritical, "SISTEMA DE SEGURIDAD"
rscar.Bookmark = str1
End If
llenarcamposcar
End Sub

Public Sub llenarcamposcar()
If rscar.BOF Then Exit Sub
If rscar.EOF Then Exit Sub
txtcod.Text = rscar.Fields("cod_car")
txtcar.Text = rscar.Fields("nom_car")
End Sub

Private Sub Command3_Click()
txtcod.SetFocus
End Sub

Private Sub Command4_Click()
Dim a As Integer
a = MsgBox("DESEA ANULAR EL CARGO MENCIONADO", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rscar.Delete
rscar.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command5_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA", vbYesNo, "SISTEMA DE SEGURIDAD")
If vbYes Then
Unload Me
frmcar.Hide
End If

End Sub

Private Sub Command6_Click()
limpiarcampos
rscli.AddNew

cod = txtcod.Text & (rscli.RecordCount)

End Sub

Private Sub Form_Load()
cod = txtcod.Text & (rscli.RecordCount)
Set rscar = New ADODB.Recordset
rscar.CursorLocation = adUseClient
sqlcar = "select *from cargos"
rscar.Open sqlcar, cn, adOpenStatic, adLockOptimistic
Set dgcar.DataSource = rscar
dgcar.Refresh

dgcar.Columns(0).Caption = "Codigo"
dgcar.Columns(1).Caption = "Cargo"
dgcar.Columns(2).Visible = False
dgcar.Columns(3).Visible = False
dgcar.Columns(4).Visible = False
dgcar.Columns(5).Visible = False
End Sub
