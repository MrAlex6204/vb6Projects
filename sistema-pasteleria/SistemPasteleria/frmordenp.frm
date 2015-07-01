VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmordenp 
   Caption         =   "Orden de Pedido"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "S"
      Height          =   300
      Left            =   3360
      TabIndex        =   12
      Top             =   3480
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "I"
      Height          =   300
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "N"
      Height          =   300
      Left            =   1800
      TabIndex        =   10
      Top             =   3480
      Width           =   600
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2355
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
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   3765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consultar por"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   -120
      Picture         =   "frmordenp.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "frmordenp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

lblfecha = Format(Date, "long Date")



End Sub

