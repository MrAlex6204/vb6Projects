VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   3135
   ClientTop       =   2970
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9885
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox pctBox 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   3000
      ScaleHeight     =   3915
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   2040
      Width           =   5415
   End
   Begin VB.HScrollBar hsbCA 
      Height          =   375
      LargeChange     =   10
      Left            =   2640
      Max             =   100
      Min             =   -100
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.HScrollBar hsbBA 
      Height          =   375
      LargeChange     =   10
      Left            =   2640
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.Frame fraEjes 
      Caption         =   "Divisiones Ejes"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2415
      Begin VB.OptionButton optSi 
         Caption         =   "SI"
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton ooptNo 
         Caption         =   "NO"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraDib 
      Caption         =   "Dibujo"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
      Begin VB.OptionButton optD2 
         Caption         =   "Mantener"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optD1 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label lblX2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblX1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblCA 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblBA 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X2/X1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X1/XR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim a, b, c As Double
Dim x1, x2, dis, xr, xi As Double
Private Sub divisiones(nx As Integer, ny As Integer)
Dim i As Integer
Dim x, xinc, y, yinc As Single
pctBox.DrawWidth = 1
xinc = 20 / (nx - 1)
x = -10
For i = 1 To nx
pctBox.Line (x, 0)-(x, -1)
x = x + xinc
Next i
yinc = 10 / (ny - 1)
y = -5
For i = 1 To ny
pctBox.Line (-1, y)-(0, y)
y = y + yinc
Next i
pctBox.DrawWidth = 2
End Sub
Private Sub cmdSalir_Click()
End
End Sub
Private Sub Form_Load()
pctBox.Scale (-10, 5)-(10, -5)
End Sub
Private Sub hsbBA_Change()
a = 1
b = hsbBA.Value / 10#
c = hsbCA.Value / 10#
lblBA.Caption = b
lblCA.Caption = c
dis = b ^ 2 - 4 * a * c
If optD2.Value = True Then 'mantener
pctBox.AutoRedraw = True
Else: 'borrar
pctBox.AutoRedraw = False
pctBox.Cls
End If
If dis > 0 Then
x1 = (-b + Sqr(dis)) / (2 * a)
x2 = (-b - Sqr(dis)) / (2 * a)
lblX1.Caption = Format(x1, "###0.000")
lblX2.Caption = Format(x2, "###0.000")
pctBox.PSet (x1, 0), vbRed
pctBox.PSet (x2, 0), vbRed
ElseIf dis = 0 Then
x1 = -b / (2 * a)
x2 = x1
lblX1.Caption = Format(x1, "###0.000")
lblX2.Caption = ""
pctBox.PSet (x1, 0), vbGreen
Else
xr = -b / (2 * a)
xi = Sqr(-dis) / (2 * a)
lblX1.Caption = Format(xr, "###0.000")
lblX2.Caption = Format(xi, "###0.000")
pctBox.PSet (xr, xi), vbBlue
pctBox.PSet (xr, -xi), vbBlue
End If

If optSi = True Then
Call divisiones(10, 5)
End If
End Sub
Private Sub hsbCA_Change()
a = 1
b = hsbBA.Value / 10#
c = hsbCA.Value / 10#
lblBA.Caption = b
lblCA.Caption = c
dis = b ^ 2 - 4 * a * c
If optD2.Value = True Then 'mantener
pctBox.AutoRedraw = True
Else: 'borrar
pctBox.AutoRedraw = False
pctBox.Cls
End If
If dis > 0 Then
x1 = (-b + Sqr(dis)) / (2 * a)
x2 = (-b - Sqr(dis)) / (2 * a)
lblX1.Caption = Format(x1, "###0.000")
lblX2.Caption = Format(x2, "###0.000")
pctBox.PSet (x1, 0), vbRed
pctBox.PSet (x2, 0), vbRed
ElseIf dis = 0 Then
x1 = -b / (2 * a)
x2 = x1
lblX1.Caption = Format(x1, "###0.000")
lblX2.Caption = ""
pctBox.PSet (x1, 0), vbGreen
Else
xr = -b / (2 * a)
xi = Sqr(-dis) / (2 * a)
lblX1.Caption = Format(xr, "###0.000")
lblX2.Caption = Format(xi, "###0.000")
pctBox.PSet (xr, xi), vbBlue
pctBox.PSet (xr, -xi), vbBlue
End If
If optSi = True Then
Call divisiones(10, 5)
End If
End Sub

Private Sub optD1_Click()
pctBox.AutoRedraw = True
pctBox.Cls
pctBox.DrawWidth = 1
pctBox.Line (-90, 0)-(90, 0), vbBlack
pctBox.Line (0, -45)-(0, 45), vbBlack
pctBox.DrawWidth = 2
End Sub

Private Sub pctBox_Paint()
pctBox.AutoRedraw = True
pctBox.Line (-90, 0)-(90, 0), vbBlack
pctBox.Line (0, -45)-(0, 45), vbBlack
pctBox.DrawWidth = 2
End Sub
