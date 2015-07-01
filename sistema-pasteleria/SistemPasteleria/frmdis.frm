VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdis 
   BackColor       =   &H8000000E&
   Caption         =   "Distritos"
   ClientHeight    =   5280
   ClientLeft      =   3675
   ClientTop       =   2835
   ClientWidth     =   6390
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6390
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   600
      Picture         =   "frmdis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   580
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   1200
      Picture         =   "frmdis.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E"
      Height          =   300
      Left            =   2400
      TabIndex        =   8
      Top             =   240
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A"
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   3600
      Picture         =   "frmdis.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   580
   End
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   3000
      Picture         =   "frmdis.frx":1156
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   580
   End
   Begin MSDataGridLib.DataGrid dgdis 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2143
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
   Begin VB.TextBox txtdis 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1980
   End
   Begin VB.TextBox txtcoddis 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distrito"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmdis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA GRABAR AL NUEVO DISTRITO", vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
copiarcampos
rsdis.Update
Else
rsdis.Cancel
End If

End Sub

Private Sub Command2_Click()

Dim b As String
str1 = rsdis.Bookmark
b = InputBox("INGRESE NOMBRE DE DISTRITO A BUSCAR")
rsdis.MoveFirst
rsdis.Find "nom_dis='" + Trim(b) + "'"
If rsdis.EOF Then
MsgBox ("NOMBRE DE DISTRITO NO EXISTE")
rsdis.Bookmark = rsdis
End If
llenarcampos



End Sub

Private Sub Command3_Click()
txtcoddis.SetFocus
End Sub

Private Sub Command4_Click()
Dim a As Integer
a = MsgBox("DESEA ANULAR DICHO DISTRITO", vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rsdis.Delete
rsdis.MoveLast
limpiarcampos
End If

End Sub

Private Sub Command5_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA", vbYesNo, "SISTEMA DE SEGURIDAD")
If vbYes Then
Unload Me
frmdis.Hide
End If
End Sub

Private Sub Command6_Click()
limpiarcampos
rsdis.AddNew


'habilitacampos

End Sub

Private Sub Form_Load()
activadis


llenarcampos

Set rsdis = New ADODB.Recordset
rsdis.CursorLocation = adUseClient
sqldis = "select *from distritos"
rsdis.Open sqldis, cn, adOpenStatic, adLockOptimistic
Set dgdis.DataSource = rsdis
dgdis.Refresh
dgdis.Columns(0).Caption = "Codigo"
dgdis.Columns(1).Caption = "Distritos"
dgdis.Columns(1).Width = 0.47 * dgdis.Width



End Sub

Public Sub llenarcampos()
If rsdis.BOF Then Exit Sub
If rsdis.EOF Then Exit Sub
txtcoddis.Text = rsdis.Fields("cod_dis")
txtdis.Text = rsdis.Fields("nom_dis")
End Sub

Public Sub copiarcampos()
rsdis.Fields("cod_dis") = txtcoddis.Text
rsdis.Fields("nom_dis") = txtdis.Text
End Sub

Public Sub limpiarcampos()
txtcoddis.Text = ""
txtdis.Text = ""
End Sub
