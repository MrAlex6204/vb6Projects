VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmordped 
   Caption         =   "Orden de Pedido"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text11 
      Height          =   300
      Left            =   7320
      TabIndex        =   32
      Top             =   5040
      Width           =   1000
   End
   Begin VB.TextBox Text10 
      Height          =   300
      Left            =   7320
      TabIndex        =   31
      Top             =   4680
      Width           =   1000
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   7320
      TabIndex        =   30
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton Command7 
      Caption         =   "S"
      Height          =   300
      Left            =   3720
      TabIndex        =   26
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "N"
      Height          =   300
      Left            =   3000
      TabIndex        =   25
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E"
      Height          =   300
      Left            =   2280
      TabIndex        =   24
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A"
      Height          =   300
      Left            =   1440
      TabIndex        =   23
      Top             =   5160
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "G"
      Height          =   300
      Left            =   720
      TabIndex        =   22
      Top             =   5160
      Width           =   600
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   360
      TabIndex        =   19
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2778
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
   Begin VB.CommandButton Command2 
      Caption         =   "SALE"
      Height          =   300
      Left            =   7440
      TabIndex        =   18
      Top             =   2280
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INGRESA"
      Height          =   300
      Left            =   7440
      TabIndex        =   17
      Top             =   1920
      Width           =   1000
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   6480
      TabIndex        =   16
      Top             =   2280
      Width           =   800
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   5760
      TabIndex        =   15
      Top             =   2280
      Width           =   555
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   5040
      TabIndex        =   14
      Top             =   2280
      Width           =   555
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   4320
      TabIndex        =   13
      Top             =   2280
      Width           =   555
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtordp 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      Height          =   195
      Left            =   6120
      TabIndex        =   29
      Top             =   5040
      Width           =   405
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pago a Cuenta"
      Height          =   195
      Left            =   6120
      TabIndex        =   28
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   6120
      TabIndex        =   27
      Top             =   4320
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Entrega"
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   4440
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      Height          =   195
      Left            =   6600
      TabIndex        =   12
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
      Height          =   195
      Left            =   5760
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.Unit"
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   2040
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   5985
      Left            =   0
      Picture         =   "frmordped.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   8640
   End
End
Attribute VB_Name = "frmordped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nroordp As String

Private Sub Form_Load()
lblfecha = Format(Date, "long Date")


activaordpedido

nroordp = "500-100-00" & (rsordpedido.RecordCount + 1)
nroordp = Right(nroordp, 11)
txtordp = nroordp
    
    
    'rsguiai.Close





End Sub

