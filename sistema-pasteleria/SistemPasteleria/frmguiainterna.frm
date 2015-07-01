VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmguiainterna 
   Caption         =   "Guia Interna"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "E"
      Height          =   300
      Left            =   2400
      TabIndex        =   22
      Top             =   3600
      Width           =   600
   End
   Begin VB.CommandButton Command6 
      Caption         =   "S"
      Height          =   300
      Left            =   3960
      TabIndex        =   21
      Top             =   3600
      Width           =   600
   End
   Begin VB.CommandButton Command5 
      Caption         =   "N"
      Height          =   300
      Left            =   3240
      TabIndex        =   20
      Top             =   3600
      Width           =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A"
      Height          =   300
      Left            =   1680
      TabIndex        =   19
      Top             =   3600
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "G"
      Height          =   300
      Left            =   960
      TabIndex        =   18
      Top             =   3600
      Width           =   600
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   6600
      TabIndex        =   16
      Top             =   3480
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
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
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTRA"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   5280
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   4680
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   6120
      TabIndex        =   17
      Top             =   3480
      Width           =   360
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
      Height          =   195
      Left            =   4680
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.Unit"
      Height          =   195
      Left            =   3960
      TabIndex        =   5
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Guia Interna"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   -120
      Picture         =   "frmguiainterna.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   7920
   End
End
Attribute VB_Name = "frmguiainterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lbolfecha = Format(Date, "long Date")

End Sub

Private Sub Image1_Click()

End Sub
