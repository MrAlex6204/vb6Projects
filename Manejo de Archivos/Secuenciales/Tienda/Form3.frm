VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmVentas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Factura"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15240
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11025
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame frameCorte 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Corte de Caja"
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
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "&Listado de Articulos"
         Default         =   -1  'True
         Height          =   345
         Left            =   3960
         TabIndex        =   38
         Top             =   1680
         Width           =   1860
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Borrar Todo"
         Height          =   345
         Left            =   6000
         TabIndex        =   37
         Top             =   1680
         Width           =   1260
      End
      Begin VB.CommandButton cmdCloseCorte 
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label16 
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
         Left            =   1440
         TabIndex        =   36
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Efectivo:"
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
         Left            =   480
         TabIndex        =   35
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblEfectivo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
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
         Left            =   1680
         TabIndex        =   34
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cajero :"
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
         Left            =   480
         TabIndex        =   33
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblNomcajero 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cajero"
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
         Left            =   1560
         TabIndex        =   32
         Top             =   720
         Width           =   705
      End
   End
   Begin VB.Frame frmListArticulos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Articulos"
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
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   7575
      Begin VB.ListBox Listart 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   7335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
   End
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
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox txtCambio 
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
         Height          =   420
         Left            =   1080
         TabIndex        =   22
         Top             =   908
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   7080
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambio: $ "
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
         Left            =   960
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pago :"
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
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Left            =   1080
         TabIndex        =   20
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total :"
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
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   8400
      TabIndex        =   5
      Top             =   5640
      Width           =   6495
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
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
            Name            =   "System"
            Size            =   9.75
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
            Name            =   "System"
            Size            =   9.75
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
            Name            =   "System"
            Size            =   9.75
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
      Top             =   7440
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ESC=Listado de Articulos  F5=Cancelar Cuenta  F6=Corte de Caja"
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
         Left            =   4440
         TabIndex        =   29
         Top             =   720
         Width           =   7020
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
         Left            =   6480
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label LabelPrecio 
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
         TabIndex        =   15
         Top             =   960
         Width           =   660
      End
      Begin VB.Label LabelDescrip 
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
         TabIndex        =   14
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "F1=Consulta   F2=Cerra  F3=Eliminar   F4=Total     "
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
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   5340
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
         Left            =   5160
         TabIndex        =   12
         Top             =   1200
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
         Left            =   6600
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblConsul2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Precio: $"
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
         Left            =   5385
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   945
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
      Height          =   4695
      Left            =   480
      TabIndex        =   16
      Top             =   2760
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8281
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
      Top             =   2280
      Width           =   5295
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
frmListArticulos.Visible = False
End Sub

Private Sub cmdCloseCorte_Click()
frameCorte.Visible = False
End Sub

Private Sub cmdOK_Click()
 If MsgBox("        ¿Desea Borrar La Cuenta del Cajero?     ", vbYesNo, "VeraSoft Development") = 6 Then
    MsgBox "Cuenta Eliminada", vbExclamation, "VeraSoft Development..."
    'Reinicia el archivo borra su contenido
    Open App.path + "\CorteCaja.Dat" For Output As #56
    Close #56
    frameCorte.Visible = False
    End If
End Sub

 Sub Command1_Click()
Label9.Visible = False
txtCambio.Enabled = True
frmTotal.Visible = False
Label4.Visible = True
End Sub

 Sub Command4_Click()

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Call Command1_Click
End If
End Sub

Private Sub Command2_Click()
frmArtVendidos.Show
End Sub

Private Sub Form_Load()
Index2 = 1
Total = 0
lblHra = Time
lblFecha = Date
LabelDescrip.Caption = "******"
LabelPrecio.Caption = "******"
frmPanel.Caption = "Le Atiende: " + CajeroOnline
End Sub

Private Sub Listart_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
frmListArticulos.Visible = False
End If
End Sub

Private Sub txtBuscar_Change()
lblConsul1.Visible = False
lblConsul2.Visible = False
lblDescrip.Visible = False
lblPrecio.Visible = False

End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Total As Double
Dim ind As Integer
ind = 1
'ESC enlista Los articulos
If KeyCode = 27 Then
frmListArticulos.Top = frmTotal.Top
frmListArticulos.Left = frmTotal.Left

Listart.Clear
Dim nArt, Desc, Pre
frmListArticulos.Visible = True

Open App.path + "\Articulos.Dat" For Input As #15
            
            Do While Not EOF(15)
            Input #15, nArt, Pre, Desc
            Listart.AddItem ind & "<-- CODE:" & nArt & " " & Desc & " $" & Pre
            ind = ind + 1
            
Loop
Close #15
Listart.SetFocus
End If

'F1 Consulta
If KeyCode = 112 Then
lblConsul1.Visible = True
lblConsul2.Visible = True
lblDescrip.Visible = True
lblPrecio.Visible = True

If BuscarArt(txtBuscar) = True Then
        lblPrecio = Precio
        lblDescrip = Descrip
    Else
    MsgBox "El Articulo No Se Encuantra !!!!", vbExclamation, "VeraSoft Development"
    End If


End If

'F2 Cerrar
If KeyCode = 113 Then

    If MsgBox("        ¿Desea Salir?        ", vbYesNo, "VeraSoft Development") = 6 Then
    'Borra el archivo temporal
    Kill App.path + "\" + CajeroOnline + ".Dat"
    Unload Me
    
    End If
    
End If

'F3 Eliminar
If KeyCode = 114 Then
    If MsgBox("        ¿Desea Eliminar Articulo?     ", vbYesNo, "VeraSoft Development") = 6 Then
      Eliminar txtBuscar, App.path + "\" + CajeroOnline + ".Dat"

        If Existe = False Then
            MsgBox "No Se Encontro Favor de Teclear Bien el Index del articulo", vbCritical, "VeraSoft Dev.."
        Else
            MsgBox "Articulo Eliminado", vbCritical, "VeraSoft Dev.."
            'Actualiza la lista
            RichText = Empty
            Open App.path + "\" + CajeroOnline + ".Dat" For Input As #15
            
            Do While Not EOF(15)
            Input #15, Index, NumArt, Descrip, Precio
            'Actualiza los Datos del RichTextBox
            
                If Len(RichText.Text) = 0 Then
                    RichText.Text = Index & "<--" & Descrip & "   $" & Precio
                Else
                    RichText.Text = RichText.Text & vbCrLf & Index & "<--" & Descrip & "   $" & Precio
                End If
        
            Loop
            Close #15

        End If
    End If
End If


'F4 Total
If KeyCode = 115 Then
frmTotal.Visible = True
txtCambio.SetFocus
Archivo = FreeFile()
Open App.path + "\" + CajeroOnline + ".Dat" For Input As #Archivo
Do While Not EOF(Archivo)
    Input #Archivo, Index2, NumArt, Descrip, Precio
        Total = Total + Precio
        
Loop
Close #Archivo
Label4 = Total

End If

'F5 Cancelacion de Cuenta
If KeyCode = 116 Then

    If MsgBox("        ¿Desea Cancelar la Cuenta?     ", vbYesNo, "VeraSoft Development") = 6 Then
    
    'Reinicia el archivo borra su contenido
    Open App.path + "\" + CajeroOnline + ".Dat" For Output As #Archivo
    Close #Archivo
    RichText = Empty
    Index2 = 1
    
    End If


End If

'F5 Corte de Caja
If KeyCode = 117 Then
frameCorte.Visible = True
lblNomcajero = CajeroOnline
frameCorte.Top = frmTotal.Top
frameCorte.Left = frmTotal.Left
'/////////////////////////////////////////
'Calcula el Corte de Caja
Dim TotalCorte As Double
TotalCorte = 0
Open App.path + "\CorteCaja.Dat" For Input As #15
    Do While Not EOF(15)
   Input #15, Index, NumArt, Descrip, Precio
   TotalCorte = TotalCorte + Precio
   
    Loop
Close #15
lblEfectivo = TotalCorte
'////////////////////////////////////////


End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
Archivo = FreeFile

'VALIDA SI EL  DATO INTRODUCIDO
'ES NUMERICO
If IsNumeric(txtBuscar.Text) = False Then
Beep
txtBuscar = Empty
Exit Sub
End If

'al presionar Enter
If KeyAscii = 13 Then

If txtBuscar = Empty Then
Beep
Exit Sub
End If

    If BuscarArt(txtBuscar) = True Then
        LabelPrecio = Precio
        LabelDescrip = Descrip
        txtBuscar = Empty
           'Agrega los Datos al RichTextBox
        If Len(RichText.Text) = 0 Then
        RichText.Text = Index2 & "<--" & Descrip & "   $" & Precio
        
        Else
        RichText.Text = RichText.Text & vbCrLf & Index2 & "<--" & Descrip & "   $" & Precio
        
        End If
        'Graba los Articulos en El archivo con el nombre del cajero
        'que esta cobrando
        Open App.path + "\" + CajeroOnline + ".Dat" For Append As #Archivo
        
        Print #Archivo, Index2
        Print #Archivo, NumArt
        Print #Archivo, Descrip
        Print #Archivo, Precio
        Close #Archivo
        
        Index2 = Index2 + 1
        
        txtBuscar = Empty
        Beep
    Else
    
        If MsgBox("        ¿Desea Agregarlo al Sistema?", vbYesNo, "VeraSoft Development") = 6 Then
    
        MDIForm1.Show
        frmArticulos.Show
        End If
    
    txtPrecio = "******"
    txtDescrip = "******"
    End If

End If




End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
'VALIDA SI EL  DATO INTRODUCIDO
'ES NUMERICO
If IsNumeric(txtCambio.Text) = False Then
Beep
txtCambio = Empty
Exit Sub
End If

If KeyAscii = 13 Then
Beep
    If txtCambio = Empty Then
    Exit Sub
    End If

   
   
   
   Label4 = Val(Label4) - Val(txtCambio)
   If Val(Label4) <= 0 Then
    Label9.Visible = True
    Label9 = "Cambio: $ " & Str(Abs(Label4))
    txtCambio.Enabled = False
    Label4.Visible = False
    Total = 0
    txtCambio = Empty
'Graba Todo lo que Vendio el Cajero al Archivo Corte de Caja
'///////////////////////////////////////////////////////////////
   Open App.path + "\CorteCaja.Dat" For Append As #15
   Open App.path + "\" + CajeroOnline + ".Dat" For Input As #Archivo
    Do While Not EOF(Archivo)
    
    Input #Archivo, Index, NumArt, Descrip, Precio
    Print #15, Index
    Print #15, NumArt
    Print #15, Descrip
    Print #15, Precio
    
    Loop
    Close #Archivo
    Close #15
'////////////////////////////////////////////////////////////////
'----------------------------------------------------------------
'Reinicia el archivo borra su contenido
'///////////////////////////////////////////////////////////////
    Archivo = FreeFile
    Open App.path + "\" + CajeroOnline + ".Dat" For Output As #Archivo
    Close #Archivo
    RichText = Empty
    Index2 = 1
'///////////////////////////////////////////////////////////////
Command1.SetFocus
   End If
   
 txtCambio = Empty
End If

End Sub
