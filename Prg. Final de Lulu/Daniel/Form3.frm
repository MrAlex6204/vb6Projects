VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Venta"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10290
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6960
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   7800
      Width           =   1095
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4200
      TabIndex        =   16
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4080
      TabIndex        =   15
      Top             =   4800
      Width           =   885
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   5880
      Width           =   855
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3600
      TabIndex        =   13
      Top             =   6360
      Width           =   1425
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3840
      TabIndex        =   8
      Top             =   6960
      Width           =   1020
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   7440
      Width           =   525
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3360
      TabIndex        =   6
      Top             =   7920
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Venta de Articulos"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   11655
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



Private Sub Command4_Click()
Text6 = Text5 * 0.15
Text7 = Text5 - Text6
End Sub


i = 0

End Sub

