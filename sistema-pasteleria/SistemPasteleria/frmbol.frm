VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmbol 
   BackColor       =   &H8000000E&
   Caption         =   "Boletas"
   ClientHeight    =   4335
   ClientLeft      =   2700
   ClientTop       =   2700
   ClientWidth     =   7725
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7725
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   6240
      TabIndex        =   19
      Top             =   3600
      Width           =   1300
   End
   Begin VB.CommandButton Command4 
      Caption         =   "S"
      Height          =   300
      Left            =   3480
      TabIndex        =   18
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "I"
      Height          =   300
      Left            =   2400
      TabIndex        =   17
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "N"
      Height          =   300
      Left            =   1200
      TabIndex        =   16
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTA"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   6000
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   5040
      TabIndex        =   20
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entre"
      Height          =   195
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Emision"
      Height          =   195
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consultar por"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "frmbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
