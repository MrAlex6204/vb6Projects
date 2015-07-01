VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmprove 
   Caption         =   "Proveedores"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
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
   Begin VB.CommandButton Command3 
      Caption         =   "CONSULTA"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S"
      Height          =   300
      Left            =   6360
      TabIndex        =   8
      Top             =   1200
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   300
      Left            =   5280
      TabIndex        =   7
      Top             =   1200
      Width           =   600
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedore"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consultar por"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   -120
      Picture         =   "frmprove.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   8040
   End
End
Attribute VB_Name = "frmprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
