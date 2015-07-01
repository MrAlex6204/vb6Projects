VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpedido 
   BackColor       =   &H8000000E&
   Caption         =   "Nota de Pedido Interno"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   3120
      Picture         =   "frmpedido.frx":0000
      TabIndex        =   26
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   3720
      Picture         =   "frmpedido.frx":030A
      TabIndex        =   25
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   4320
      Picture         =   "frmpedido.frx":0614
      TabIndex        =   24
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtvalores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   20
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txthora 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid GrDatArticulos 
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   3000
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1720
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
            LCID            =   10250
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
            LCID            =   10250
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
   Begin VB.TextBox txtfecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtpre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   6855
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtsubtotal 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtcproducto 
         Height          =   315
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtcantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin MSDataListLib.DataList dtllinea 
         Height          =   840
         Left            =   1080
         TabIndex        =   23
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1482
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Sub-Total"
         Height          =   195
         Left            =   4200
         TabIndex        =   30
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Linea"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Precio"
         Height          =   195
         Left            =   4800
         TabIndex        =   19
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Producto"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6855
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   27
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   5
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Encargado"
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Señores"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Nro Pedido Interno"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      Begin VB.TextBox txtpedi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Impòrte"
      Height          =   195
      Left            =   5280
      TabIndex        =   18
      Top             =   4320
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ventas a Produccion"
      Height          =   195
      Left            =   1560
      TabIndex        =   15
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "PANADERIA PASTELERIA ALISSON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   14
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Hora "
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Fecha Entrega"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1050
   End
End
Attribute VB_Name = "frmpedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nrosec As String
Dim cod_pro As String


Private Sub CmdMantenimiento_Click(Index As Integer)
Select Case Index
        Case 0 'Nuevo
            Unload Me
            frmpedido.Show
        Case 1 'Guadar
           activapedidoint
            '***************************************************************************
            rspedint.AddNew
            rspedint.Fields("numpedidoint") = txtpedi.Text
            rspedint.Fields("Fecha") = Date
            rspedint.Fields("señores") = txtdatos(1)
           rspedint.Fields("encargado") = txtdatos(0)
            'rsfac.Fields ("guia_rem")
            'rsfac.Fields ("sub_tot")
            'rsfac.Fields ("IGV")
            rspedint.Fields("importe") = txtimporte
            rspedint.Fields("fecha_ent") = txtfecha.Text
            rspedint.Fields("hora") = txthora.Text
            
            'rsfac.Fields ("DNI_per")
            'rsfac.Fields ("nom_per")
            'rsfac.Fields("cod_est") = Datacboest.Text
            'rsfac.Fields("cod_for") = Datacbofor.Text
            
            rspedint.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
            activadpedido
            
            Do While Not RsTemporal.EOF
                
                rsdpedido.AddNew
                rsdpedido.Fields("numpedidoint") = nrosec
                rsdpedido.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                rsdpedido.Fields("des_pro") = RsTemporal.Fields("Nom_Art")
                rsdpedido.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                rsdpedido.Fields("precio") = RsTemporal.Fields("Precio")
                rsdpedido.Fields("Importe") = RsTemporal.Fields("sub_total")
                rsdpedido.Update
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
            datafacturas.Caption = "Nota Pedido Interno"
            datafacturas.Show
    End Select
End Sub




Private Sub Datacbopro_Click(Area As Integer)
rspro.MoveFirst
des_pro = Datacbopro.Text
rspro.Find "des_pro='" + Trim(Datacbopro.Text) + "'"
If rspro.EOF Then
Else
End If
End Sub

Private Sub cmdopciones_Click(Index As Integer)
Select Case Index
        Case 0 'Aceptar Item
         If Val(txtcantidad) = 0 Then MsgBox "debe ingresar una cantidad mayor a cero", vbCritical, "Sistema de ACONT": Exit Sub
            With RsTemporal
                If Not .BOF Then
                    .MoveFirst
                    .Find "cod_pro='" + Trim(dtcproducto.BoundText) + "'"
                    If .EOF = False Then
                        GrDatArticulos.Columns(2) = Val(txtcantidad)
                        GrDatArticulos.Columns(5) = Val(txtsubtotal)
                        
                        
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
                    frameDetalle.Enabled = False
                    RsTemporal.Close
                    ActivaTemporal
                End If
            Else
                cmdopciones(1).Enabled = False
                txtvalores(2) = FormatCurrency(txtvalores(0), 2)
               
                
            End If
            Call GrDatArticulos_Change
    End Select
End Sub

Private Sub dtcproducto_Click(Area As Integer)
rspro.MoveFirst
rspro.Find "des_pro='" + Trim(dtcproducto.Text) + "'"
If rspro.EOF Then
Else

        txtpre = (rspro.Fields("Pre_ven"))
        txtcantidad.Enabled = True
        txtcantidad = Val(0)
        txtcantidad.SetFocus

End If
End Sub

Private Sub dtllinea_Click()
If dtllinea <> Empty Then
        dtcproducto.Enabled = True
       ' Llena el DATACOMBO (dtcproducto) con los nombre de los articulos según su categoría
        
        Set rspro = New ADODB.Recordset
        rspro.Open "select *from productos where cod_lin='" + Trim(dtllinea.BoundText) + "'", cn, adOpenStatic, adLockOptimistic
        
        dtcproducto.Text = ""
         txtpre = ""
        txtcantidad = Val(0)
      cmdopciones(0).Enabled = False
        txtcantidad.Enabled = False
     
       
       
        dtcproducto.BoundColumn = "cod_pro"
        dtcproducto.ListField = "des_pro"
        Set dtcproducto.RowSource = rspro
    End If
End Sub

Private Sub Form_Load()
activapedidoint
    nrosec = "00-300-000" & (rspedint.RecordCount + 1)
    nrosec = Right(nrosec, 11)
    txtpedi = nrosec
    rspedint.Close
    
      Call ActivaTemporal
    
    
  
     
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



Private Sub Txtcantidad_Change()
txtsubtotal = (Val(txtcantidad) * (Val(txtpre)))
  txtsubtotal = Format(txtsubtotal, "##0.00")
  txtvalores(2) = Format(txtsubtotal, "##0.00")
    
    cmdopciones(0).Enabled = True
End Sub

Private Sub txtdatos_Change(Index As Integer)
Select Case Index
        Case 0
            If Len(txtdatos(0)) = 15 Then
               ' activacli
                'rscli.MoveFirst
                'rscli.Find "DNI_cli ='" + Trim(txtdatos(0)) + "'"
                'If Not rscli.EOF Then
                      'txtdatos(1) = rscli.Fields("cod_dis")
                    'txtdatos(2) = rscli.Fields("nom_cli")
                    'txtdatos(3) = rscli.Fields("dir_cli")
                    'frameCabecera.Enabled = True
                'End If
                If Not txtdatos(1) = "" Then
                    txtdatos(1).Enabled = False
                    'txtdatos(2).Enabled = False
                    'txtdatos(3).Enabled = False
                Else
                    txtdatos(1).Enabled = True
                    'txtdatos(2).Enabled = True
                    'txtdatos(3).Enabled = True
                    txtdatos(1).SetFocus
                    
                    'frameCabecera.Enabled = True
                End If
            Else
                'txtdatos(1) = ""
                'txtdatos(2) = ""
                'txtdatos(3) = ""
                
            End If
    End Select
If txtdatos(1) <> "" Then
'       Llena el DATALIST (dtlCategoria) con los nombre de categoría
       activalin
        dtllinea.BoundColumn = "cod_lin"
        dtllinea.ListField = "nom_lin"
        Set dtllinea.RowSource = rslin
    End If
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
txtdatos(0).SetFocus
End If
End Sub
Sub ActivaTemporal()
    'CREANDO RECORDSET TEMPORAL****************
    Set RsTemporal = New ADODB.Recordset
    RsTemporal.CursorType = adOpenStatic
    RsTemporal.Fields.Append "IdEnProf", adVarChar, 10, adFldIsNullable
    RsTemporal.Fields.Append "cod_art", adVarChar, 12, adFldIsNullable
    RsTemporal.Fields.Append "Nom_Art", adVarChar, 250, adFldIsNullable
    RsTemporal.Fields.Append "Precio", adSingle, 10.2, adFldIsNullable
    RsTemporal.Fields.Append "cantidad", adInteger, adFldIsNullable
    RsTemporal.Fields.Append "Sub_total", adSingle, 10.2, adFldIsNullable
    RsTemporal.Open
    '*****************************************************
     Set GrDatArticulos.DataSource = RsTemporal
     GrDatArticulos.Columns(0).Visible = False
    GrDatArticulos.Columns(1).Visible = False
    
    GrDatArticulos.Columns(1).Caption = "ITEMS"
    GrDatArticulos.Columns(2).Caption = "DESCRIPCION"
    GrDatArticulos.Columns(3).Caption = "PRECIO"
    GrDatArticulos.Columns(4).Caption = "CANTIDAD"
    GrDatArticulos.Columns(5).Caption = "SUB TOTAL"
    GrDatArticulos.Columns(1).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(2).Width = 0.4 * GrDatArticulos.Width
    GrDatArticulos.Columns(3).Width = 0.14 * GrDatArticulos.Width
    GrDatArticulos.Columns(4).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(5).Width = 0.2 * GrDatArticulos.Width
GrDatArticulos.Columns(3).NumberFormat = "##0.00"
GrDatArticulos.Columns(5).NumberFormat = "##0.00"
GrDatArticulos.Columns(3).Alignment = dbgRight
GrDatArticulos.Columns(4).Alignment = dbgRight
GrDatArticulos.Columns(5).Alignment = dbgRight

End Sub
Sub Graba_Temporal()
    RsTemporal.AddNew
    RsTemporal.Fields(0) = Trim(txtnumero)
    RsTemporal.Fields(1) = Trim(dtcproducto.BoundText)
    RsTemporal.Fields(2) = Trim(dtcproducto.Text)
    RsTemporal.Fields(3) = Trim(txtpre)
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
        'txtvalores(2) = SubTotal
        'txtvalores(2) = Format(txtvalores(2), "##0.00")
     'txtigv = Val(TxtValores(2)) * 0.19
     'txtigv = Format(txtigv, "##0.00")
     
     'txttot.Text = Val(TxtValores(2)) + Val(txtigv.Text)
     'txttot.Text = Format(txttot, "##0.00")
     
    End If
End Sub
