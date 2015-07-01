VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmgrafico 
   BackColor       =   &H8000000E&
   Caption         =   "Cliente "
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin MSChart20Lib.MSChart grafico 
      Height          =   5775
      Left            =   840
      OleObjectBlob   =   "frmgrafico.frx":0000
      TabIndex        =   1
      Top             =   1800
      Width           =   6855
   End
   Begin MSDataGridLib.DataGrid dggra 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
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
            Type            =   1
            Format          =   "d-mmm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
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
End
Attribute VB_Name = "frmgrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim mes As String

Private Sub Command1_Click()
Unload Me
frmgrafico.Hide
End Sub

Private Sub Form_Activate()
activagra
grafico.Visible = True
grafico.Title = "Ingresos Mensuales"
grafico.chartType = VtChChartType2dCombination
grafico.ColumnCount = rsgra.RecordCount
grafico.RowCount = 1
grafico.RowLabel = "Mes"
activames
Set dggra.DataSource = rsgra
dggra.Columns(0).Caption = "Mes"
dggra.Columns(1).Caption = "Total"
dggra.Columns(1).NumberFormat = "###.0.00"
dggra.Columns(1).Alignment = dbgRight
dggra.Columns(1).Width = 1000



End Sub
Public Sub activames()
rsgra.MoveFirst
For i = 1 To rsgra.RecordCount
Select Case rsgra.Fields(0)
Case 1
mes = "Enero"
Case 2
mes = "Febrero"
Case 3
mes = "Marzo"
Case 4
mes = "Abril"
Case 5
mes = "Mayo"
Case 6
mes = "Junio"
Case 7
mes = "Julio"
Case 8
mes = "Agosto"
Case 9
mes = "Setiembre"
Case 10
mes = "Octubre"
Case 11
mes = "Noviembre"
Case 12
mes = "Diciembre"
End Select
grafico.Column = i
grafico.ColumnLabel = mes
'grafico.Data = rsgra.Fields(1)
rsgra.MoveNext
Next





End Sub
