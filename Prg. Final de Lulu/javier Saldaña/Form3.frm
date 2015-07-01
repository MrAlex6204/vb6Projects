VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Factura"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10290
   LinkTopic       =   "Form3"
   ScaleHeight     =   11085
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   11055
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Calcular"
         Height          =   495
         Left            =   3960
         TabIndex        =   13
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   1920
         TabIndex        =   18
         Top             =   3360
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   1080
         TabIndex        =   17
         Top             =   3840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Id Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   555
      Left            =   7200
      TabIndex        =   5
      Top             =   1800
      Width           =   1890
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
If i > 10 Then
MsgBox "No hay Espacio ", vbCritical
Exit Sub
Else
 
 VentasVendedores(i) = Text4
 
VentasArticulos(i) = Text3

 i = i + 1
End If


Text1 = Empty
Text2 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty


End Sub

Private Sub Command2_Click()
Text1 = Empty
Text2 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty
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

