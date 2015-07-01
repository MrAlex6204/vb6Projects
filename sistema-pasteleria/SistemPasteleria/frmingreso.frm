VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmingreso 
   BackColor       =   &H8000000E&
   Caption         =   "Documento Entrada"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtobser 
      Height          =   855
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   46
      Top             =   5160
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo datacboest 
      Height          =   315
      Left            =   1560
      TabIndex        =   44
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtpersona 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   42
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdborrar 
      Height          =   300
      Left            =   6960
      TabIndex        =   40
      Top             =   3360
      Width           =   400
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   360
      Picture         =   "frmingreso.frx":0000
      TabIndex        =   39
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   960
      Picture         =   "frmingreso.frx":030A
      TabIndex        =   38
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   1560
      Picture         =   "frmingreso.frx":0614
      TabIndex        =   37
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton cmdopciones 
      Height          =   300
      Index           =   1
      Left            =   6480
      TabIndex        =   36
      Top             =   3360
      Width           =   400
   End
   Begin MSDataListLib.DataCombo datacbodis 
      Height          =   315
      Left            =   4680
      TabIndex        =   35
      Top             =   3000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox txtdir 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   34
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txts 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   31
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtemp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   29
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Num Documento Entrada"
      Height          =   615
      Left            =   5400
      TabIndex        =   26
      Top             =   360
      Width           =   2055
      Begin VB.TextBox txtdoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtigv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   25
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtvalores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   6360
      TabIndex        =   24
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txttot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   23
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdopciones 
      Height          =   300
      Index           =   0
      Left            =   6000
      TabIndex        =   19
      Top             =   3360
      Width           =   400
   End
   Begin MSDataGridLib.DataGrid GrDatArticulos 
      Height          =   1095
      Left            =   360
      TabIndex        =   18
      Top             =   3960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1931
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
   Begin VB.TextBox TxtSubTotal 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5280
      TabIndex        =   13
      Top             =   3600
      Width           =   600
   End
   Begin VB.TextBox txtprecio 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4080
      TabIndex        =   12
      Top             =   3600
      Width           =   480
   End
   Begin VB.TextBox txtcantidad 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4680
      TabIndex        =   11
      Top             =   3600
      Width           =   480
   End
   Begin VB.TextBox txtencargado 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Width           =   2400
   End
   Begin VB.TextBox txtproducto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox txtfechado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtnumdo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txttip 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Observacion"
      Height          =   195
      Left            =   3600
      TabIndex        =   45
      Top             =   5160
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Estado"
      Height          =   195
      Left            =   360
      TabIndex        =   43
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Persona Recibe"
      Height          =   195
      Left            =   360
      TabIndex        =   41
      Top             =   5160
      Width           =   1140
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Distrito"
      Height          =   195
      Left            =   3840
      TabIndex        =   33
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Direccion"
      Height          =   195
      Left            =   360
      TabIndex        =   32
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Encargado"
      Height          =   195
      Left            =   3840
      TabIndex        =   30
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Tipo"
      Height          =   195
      Left            =   360
      TabIndex        =   28
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   5400
      TabIndex        =   22
      Top             =   6840
      Width           =   345
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   5400
      TabIndex        =   21
      Top             =   6480
      Width           =   270
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   5400
      TabIndex        =   20
      Top             =   6120
      Width           =   675
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   5280
      TabIndex        =   17
      Top             =   3360
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   4680
      TabIndex        =   16
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   4080
      TabIndex        =   15
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3600
      TabIndex        =   14
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   375
      TabIndex        =   2
      Top             =   3360
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3480
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nrosec As String

Private Sub cmdborrar_Click()
txtproducto.Text = ""
txts.Text = ""
TxtPrecio.Text = ""
txtcantidad.Text = ""
txtsubtotal.Text = ""
End Sub

Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
        Case 0 'Nuevo
            Unload Me
            frmingreso.Show
        Case 1 'Guadar
            activado
            '***************************************************************************
            rsdo.AddNew
            rsdo.Fields("num_docent") = txtdoc.Text
            rsdo.Fields("Fecha_emi") = Date
            rsdo.Fields("tipo_doc") = txttip.Text
            rsdo.Fields("descrip_doc") = txtdes.Text
            rsdo.Fields("fecha_doc") = txtfechado.Text
            
            rsdo.Fields("subtotal") = txtvalores(2)
            rsdo.Fields("IGV") = txtigv.Text
            rsdo.Fields("total") = txttot.Text
            rsdo.Fields("cod_perrec") = txtpersona.Text
            rsdo.Fields("cod_est") = Datacboest.Text
            rsdo.Fields("observacion") = txtobser.Text
            
            rsdo.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
          activaddo
            
            Do While Not RsTemporal.EOF
                
                rsddo.AddNew
                rsddo.Fields("num_docent") = nrosec
                rsddo.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                rsddo.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
                rsddo.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                'rsddo.Fields("precio") = RsTemporal.Fields("Precio")
                rsddo.Fields("importe") = RsTemporal.Fields("sub_total")
                rsddo.Update
                RsTemporal.MoveNext
            Loop
            CmdMantenimiento(0).Enabled = True
            CmdMantenimiento(1).Enabled = False
            CmdMantenimiento(2).Enabled = True
            GrDatArticulos.Enabled = False
            'framePrincipal.Enabled = False
            'frameCabecera.Enabled = False
            '********************************
        Case 2 'Previo
            If DataEnvironment1.rscmdfacturas.State = adStateOpen Then
                DataEnvironment1.rscmdfacturas.Close
            End If
                
            DataEnvironment1.cmdfacturas (Trim(TxtNumeroProforma))
            Set datafacturas.DataSource = DataEnvironment1.rscmdfacturas
            datafacturas.Caption = "Facturas"
            datafacturas.Show
    End Select
End Sub

Private Sub cmdopciones_Click(Index As Integer)
Select Case Index
        Case 0 'Aceptar Item
         If Val(txtcantidad) = 0 Then MsgBox "debe ingresar una cantidad mayor a cero", vbCritical, "Sistema de ACONT": Exit Sub
            With RsTemporal
                If Not .BOF Then
                    .MoveFirst
                    .Find "cod_art='" + Trim(txtproducto.Text) + "'"
                    If .EOF = False Then
                        GrDatArticulos.Columns(2) = Val(txtcantidad)
                        GrDatArticulos.Columns(3) = Val(txtsubtotal)
                        RsTemporal.Update
                        GrDatArticulos.Refresh
                    Else
                        Graba_Temporal
                    End If
               Else
                    Graba_Temporal
                End If
            End With
            cmdopciones(1).Enabled = True
            ' *************************************
            Call Calculos
            GrDatArticulos.Refresh
            Call GrDatArticulos_Change
            'frameDetalle.Enabled = True
        Case 1 'ELIMINAR ITEM
            If RsTemporal.RecordCount > 0 Then
                If RsTemporal.BOF Or RsTemporal.EOF Then MsgBox "Debe seleccionar de la grilla el Titulo a eliminar", vbCritical, "ACONT": Exit Sub
                RsTemporal.Delete
                RsTemporal.MoveNext
                Call Calculos
                If RsTemporal.RecordCount = 0 Then
                    'frameDetalle.Enabled = False
                    RsTemporal.Close
                    ActivaTemporal
                End If
            Else
                cmdopciones(1).Enabled = False
                txtvalores(0) = FormatCurrency(txtvalores(0), 2)
            End If
            Call GrDatArticulos_Change
    End Select
End Sub

Private Sub datacbodis_Click(Area As Integer)
rsdis.MoveFirst
nom_dis = datacbodis.Text
rsdis.Find "nom_dis='" + Trim(datacbodis.Text) + "'"
If rsdis.EOF Then
Else
End If
End Sub


Private Sub datacboest_Click(Area As Integer)
rsest.MoveFirst
nom_est = Datacboest.Text
rsest.Find "nom_est='" + Trim(Datacboest.Text) + "'"
If rsest.EOF Then
Else
End If
End Sub

Private Sub Form_Load()
activado
nrosec = "00-900-000" & (rsdo.RecordCount + 1)
    nrosec = Right(nrosec, 11)
    txtdoc = nrosec
    rsdo.Close
Call ActivaTemporal

activadis
datacbodis.ListField = "nom_dis"
Set datacbodis.RowSource = rsdis

activaest
Datacboest.ListField = "nom_est"
Set Datacboest.RowSource = rsest


End Sub

Private Sub GrDatArticulos_Change()
If GrDatArticulos.ApproxCount = 0 Then
        CmdMantenimiento(1).Enabled = False
        cmdopciones(1).Enabled = False
    Else
        CmdMantenimiento(1).Enabled = True
        cmdopciones(1).Enabled = True
    End If
    LblContador = "se encuentran profomando " + Str(GrDatArticulos.ApproxCount) + " articulos"
End Sub
Private Sub txtcantidad_Change()
txtsubtotal = (Val(txtcantidad) * (Val(TxtPrecio)))
  txtsubtotal = Format(txtsubtotal, "##0.00")
    cmdopciones(0).Enabled = True
End Sub
Sub ActivaTemporal()
    'CREANDO RECORDSET TEMPORAL****************
    Set RsTemporal = New ADODB.Recordset
    RsTemporal.CursorType = adOpenStatic
    RsTemporal.Fields.Append "IdEnProf", adVarChar, 10, adFldIsNullable
    RsTemporal.Fields.Append "cod_art", adVarChar, 12, adFldIsNullable
    RsTemporal.Fields.Append "Nom_Art", adVarChar, 250, adFldIsNullable
    RsTemporal.Fields.Append "Precio", adDouble, 10.2, adFldIsNullable
    RsTemporal.Fields.Append "cantidad", adInteger, adFldIsNullable
    RsTemporal.Fields.Append "Sub_total", adDouble, 10.2, adFldIsNullable
    RsTemporal.Open
    '*****************************************************
     Set GrDatArticulos.DataSource = RsTemporal
     GrDatArticulos.Columns(0).Visible = False
    GrDatArticulos.Columns(1).Visible = False
    
    GrDatArticulos.Columns(1).Caption = "ITEMS"
    GrDatArticulos.Columns(2).Caption = "Descripcion"
    GrDatArticulos.Columns(3).Caption = "Precio"
    GrDatArticulos.Columns(4).Caption = "Cantidad"
    GrDatArticulos.Columns(5).Caption = "Sub Total"
    GrDatArticulos.Columns(1).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(2).Width = 0.4 * GrDatArticulos.Width
    GrDatArticulos.Columns(3).Width = 0.13 * GrDatArticulos.Width
    GrDatArticulos.Columns(4).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(5).Width = 0.2 * GrDatArticulos.Width
GrDatArticulos.Columns(3).NumberFormat = "##,##0.00"
GrDatArticulos.Columns(5).NumberFormat = "##,##0.00"


GrDatArticulos.Columns(3).Alignment = dbgRight
GrDatArticulos.Columns(4).Alignment = dbgRight
GrDatArticulos.Columns(5).Alignment = dbgRight

End Sub
Sub Graba_Temporal()
    RsTemporal.AddNew
    RsTemporal.Fields(0) = Trim(txtnumero)
    RsTemporal.Fields(1) = Trim(txtproducto.Text)
    RsTemporal.Fields(2) = Trim(txtproducto.Text)
    RsTemporal.Fields(3) = Trim(TxtPrecio)
    RsTemporal.Fields(4) = Trim(txtcantidad)
    RsTemporal.Fields(5) = Trim(txtsubtotal)
    RsTemporal.Update
End Sub
Sub Calculos()
    TxtPrecio = FormatCurrency(TxtPrecio, 2)
    If RsTemporal.RecordCount > 0 Then
        Dim SubTotal As Double
        RsTemporal.MoveFirst
        Do
            SubTotal = SubTotal + RsTemporal.Fields("Sub_Total")
            RsTemporal.MoveNext
        Loop While Not RsTemporal.EOF
        txtvalores(2) = SubTotal
        txtvalores(2) = Format(txtvalores(2), "##0.00")
     txtigv = Val(txtvalores(2)) * 0.19
     txtigv = Format(txtigv, "##0.00")
     
     txttot.Text = Val(txtvalores(2)) + Val(txtigv.Text)
     txttot.Text = Format(txttot, "##0.00")
     
    End If
End Sub

