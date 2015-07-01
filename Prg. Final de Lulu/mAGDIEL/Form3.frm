VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Factura"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10290
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   720
      TabIndex        =   10
      Top             =   6600
      Width           =   7935
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Calcular"
         Height          =   495
         Left            =   3720
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Total Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "I.V.A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "SubTotal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   7935
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nº Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Articulo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nº Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2280
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   6870
      Left            =   8760
      Picture         =   "Form3.frx":0000
      Top             =   2520
      Width           =   6420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Venta de Articulos"
      BeginProperty Font 
         Name            =   "Asimov"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Width           =   3675
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Command1_Click()

 
 VentasVendedores(i) = Text4
 
VentasArticulos(i) = Text3

 i = i + 1






End Sub



Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Command4_Click()
Text6 = Text5 * 0.15
Text7 = Text5 - Text6
End Sub

Private Sub Form_Load()

i = 0

End Sub

