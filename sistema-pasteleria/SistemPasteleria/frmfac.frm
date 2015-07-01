VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmfac 
   BackColor       =   &H80000009&
   Caption         =   "Facturas"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   7200
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   7200
      TabIndex        =   20
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   7200
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "S"
      Height          =   300
      Left            =   3000
      TabIndex        =   18
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "I"
      Height          =   300
      Left            =   2040
      TabIndex        =   17
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "N"
      Height          =   300
      Left            =   1080
      TabIndex        =   16
      Top             =   3840
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTA"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   8295
      _ExtentX        =   14631
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
      Left            =   7320
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1440
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   6240
      TabIndex        =   24
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      Height          =   195
      Left            =   6240
      TabIndex        =   23
      Top             =   4080
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Left            =   6240
      TabIndex        =   22
      Top             =   3720
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entre"
      Height          =   195
      Left            =   7560
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas Emitas"
      Height          =   195
      Left            =   5520
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   930
   End
End
Attribute VB_Name = "frmfac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
