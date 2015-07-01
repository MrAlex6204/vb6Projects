VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmmar 
   BackColor       =   &H8000000E&
   Caption         =   "Marcas"
   ClientHeight    =   3060
   ClientLeft      =   3735
   ClientTop       =   930
   ClientWidth     =   5280
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5280
   Begin VB.CommandButton Command1 
      Height          =   610
      Left            =   840
      Picture         =   "frmmar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command2 
      Height          =   610
      Left            =   1440
      Picture         =   "frmmar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A"
      Height          =   300
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E"
      Height          =   300
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Height          =   610
      Left            =   3240
      Picture         =   "frmmar.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   580
   End
   Begin VB.CommandButton Command6 
      Height          =   610
      Left            =   3840
      Picture         =   "frmmar.frx":1156
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   580
   End
   Begin MSDataGridLib.DataGrid dgmar 
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   1680
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
   Begin VB.TextBox txtmar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   2220
   End
   Begin VB.TextBox txtcodmar 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   450
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
Attribute VB_Name = "frmmar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim g As Integer
g = MsgBox("DESEA REGISTRAR LA MARCA", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If g = vbYes Then
'-------------------------
copiarcampos
rsmar.Update
'--------------------
Else
rsmar.Cancel
End If
End Sub

Private Sub Command2_Click()
Dim b As String
str1 = rsmar.Bookmark
b = InputBox("INGRESE NOMBRE DE LA MARCA A BUSCAR", "SISTEMA DE SEGURIDAD")
rsmar.MoveFirst
rsmar.Find "nom_mar='" + Trim(b) + "'"
If rsmar.EOF Then
MsgBox "NOMBRE DE LA MARCA NO EXISTE", vbCritical, "SISTEMA DE SEGURIDAD"
rsmar.Bookmark = srt1
End If
llenarcampos
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("DESEA ANULAR LA MARCA", vbQuestion + vbYesNo, "SISTEMA DE SEGURIDAD")
If a = vbYes Then
rsmar.Delete
rsmar.MoveLast
limpiarcampos
End If
End Sub

Private Sub Command4_Click()
txtcodmar.SetFocus
End Sub

Private Sub Command5_Click()
limpiarcampos
'___________________

rsmar.AddNew
'___________________


txtcodmar.SetFocus
End Sub

Private Sub Command6_Click()
Dim X As Integer
X = MsgBox("DESEA SALIR DEL SISTEMA", vbYesNo, "SISTEMA DE SEGURIDAD")
If X = vbYes Then
Unload Me
End If
frmmar.Hide
End Sub

Private Sub Form_Load()
activamar
llenarcampos

Set rsmar = New ADODB.Recordset
rsmar.CursorLocation = adUseClient
sqlmar = "select *from marcas"
rsmar.Open sqlmar, cn, adOpenStatic, adLockOptimistic
Set dgmar.DataSource = rsmar
dgmar.Refresh
dgmar.Columns(0).Caption = "Codigo"
dgmar.Columns(1).Caption = "Marcas"
dgmar.Columns(1).Width = 0.47 * dgmar.Width

End Sub


Public Sub llenarcampos()
If rsmar.BOF Then Exit Sub
If rsmar.EOF Then Exit Sub
txtcodmar.Text = rsmar.Fields("cod_mar")
txtmar.Text = rsmar.Fields("nom_mar")
End Sub

Public Sub copiarcampos()
'________________________________________________
rsmar.Fields("cod_mar") = txtcodmar.Text
rsmar.Fields("nom_mar") = txtmar.Text
'_______________________________________________
End Sub

Public Sub limpiarcampos()
txtcodmar.Text = ""
txtmar.Text = ""
End Sub

