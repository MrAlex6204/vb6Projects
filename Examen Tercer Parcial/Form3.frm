VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Factura"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15360
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   2175
      Left            =   600
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   8400
      TabIndex        =   5
      Top             =   6000
      Width           =   6495
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "@MS Mincho"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "@MS Mincho"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblHra 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "@MS Mincho"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "@MS Mincho"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame frmPanel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Le atiende:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   7920
      Width           =   11655
      Begin VB.TextBox txtBuscar 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label txtPrecio 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   270
         Left            =   2040
         TabIndex        =   17
         Top             =   960
         Width           =   660
      End
      Begin VB.Label txtDescrip 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   270
         Left            =   2040
         TabIndex        =   16
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "F1=Consulta      Enter=Aceptar      F4=Total     F2=Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label lblSigno 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   6840
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblConsul1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5280
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "descrip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   6840
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblPrecio 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   7000
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblConsul2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Precio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5880
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Num:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1305
      End
   End
   Begin RichTextLib.RichTextBox RichText 
      Height          =   5055
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   2985
      Left            =   2640
      Picture         =   "Form3.frx":0082
      Top             =   -360
      Width           =   11640
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   8520
      Picture         =   "Form3.frx":7126C
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   5295
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmTotal.Visible = False
End Sub

 Sub Command4_Click()

End Sub

Private Sub Form_Load()
lblHra = Time
lblFecha = Date
txtDescrip.Caption = "******"
txtPrecio.Caption = "******"
frmPanel.Caption = "Le Atiende: " + Datos.NomCajero
End Sub




Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label9_Click()
End Sub

Private Sub txtBuscar_Change()
lblConsul1.Visible = False
lblConsul2.Visible = False
lblDescrip.Visible = False
lblPrecio.Visible = False
lblSigno.Visible = False
End Sub



Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = 113 Then
If MsgBox("¿Desea Cerrar?", vbYesNo, "Verasoft Development") = 6 Then
Unload Me
End If

End If

If KeyCode = 115 Then
frmTotal.Visible = True

End If

If KeyCode = 112 Then
lblConsul1.Visible = True
lblConsul2.Visible = True
lblDescrip.Visible = True
lblPrecio.Visible = True
lblSigno.Visible = True

Datos.AbrirBase ("Select * from ListArt")

Datos.Buscar txtBuscar, "NumArt"

If Datos.Encontrado = False Then
Datos.CerrarBase 'Cierra la base de Datos
Exit Sub
Else

lblDescrip = BaseDatosOpen.Fields("Descrip")
lblPrecio = BaseDatosOpen.Fields("Precio")

End If
Datos.CerrarBase
End If

End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

Datos.AbrirBase ("Select * from ListArt") 'Abre la tabla de la BaseDeDatos
Datos.Buscar txtBuscar, "NumArt" 'Funcion sirve Para Buscar


'encaso de q no se encontrara el producto
If Datos.Encontrado = False Then
Datos.CerrarBase 'Cierra la base de Datos
Exit Sub
Else

'Encasa de que si se encontrara

txtDescrip = BaseDatosOpen.Fields("Descrip")
txtPrecio = BaseDatosOpen.Fields("Precio")

End If


Datos.CerrarBase
If Len(RichText.Text) = 0 Then
 RichText.Text = txtDescrip + "   $" + txtPrecio
 Else
 RichText.Text = RichText.Text + vbCrLf + txtDescrip + "   $" + txtPrecio
 End If
txtBuscar = Empty

End If

End Sub
