VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpagoc 
   BackColor       =   &H8000000E&
   Caption         =   "Pagos a Cuentas"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "S"
      Height          =   300
      Left            =   3840
      TabIndex        =   18
      Top             =   4680
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I"
      Height          =   300
      Left            =   2880
      TabIndex        =   17
      Top             =   4680
      Width           =   600
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2143
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
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3240
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   840
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo Datacbofecha 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin MSDataListLib.DataCombo Datacbopago 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      Height          =   195
      Left            =   4800
      TabIndex        =   15
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Cuenta"
      Height          =   195
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden de Pedido"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   765
   End
End
Attribute VB_Name = "frmpagoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
