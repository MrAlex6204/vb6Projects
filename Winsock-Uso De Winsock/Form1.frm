VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Texto /Ficheros Recibidos"
      Height          =   1815
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   9375
      Begin VB.TextBox Text10 
         Height          =   615
         Left            =   840
         TabIndex        =   26
         Text            =   "Text10"
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fichero Recidido"
         Height          =   195
         Left            =   4920
         TabIndex        =   27
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "Texto Recibido"
         Height          =   435
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instancias Activas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   9480
      TabIndex        =   20
      Top             =   240
      Width           =   4815
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Text            =   "Text9"
         Top             =   1200
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   "Ultima Instancia Eliminada"
         Height          =   795
         Left            =   3360
         TabIndex        =   22
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maquina Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   7920
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Escuchar"
         Height          =   375
         Left            =   7080
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Puerto"
         Height          =   195
         Left            =   7320
         TabIndex        =   18
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "IP"
         Height          =   195
         Left            =   4920
         TabIndex        =   17
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ultima Maquina Remota Conectada"
         Height          =   195
         Left            =   4920
         TabIndex        =   15
         Top             =   960
         Width           =   2520
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Resquest ID"
         Height          =   195
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Maquina Local"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP Local"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Puerto Local"
         Height          =   195
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label15_Click()

End Sub
