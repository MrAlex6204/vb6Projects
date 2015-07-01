VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_comp 
   BackColor       =   &H80000009&
   Caption         =   "Orden de Compra"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtvalores 
      Height          =   285
      Index           =   2
      Left            =   6840
      TabIndex        =   27
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   3120
      Picture         =   "frm_comp.frx":0000
      TabIndex        =   26
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   3720
      Picture         =   "frm_comp.frx":030A
      TabIndex        =   25
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   4320
      Picture         =   "frm_comp.frx":0614
      TabIndex        =   24
      Top             =   4080
      Width           =   615
   End
   Begin VB.Frame frameCabecera 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   7695
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtstock 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Txtcantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox TxtPrecio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox TxtSubTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtcproducto 
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataList dtllinea 
         Height          =   645
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1138
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Stock"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4920
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "Linea de Productos"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Lista de productos"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   2640
         TabIndex        =   18
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cant"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   5760
         TabIndex        =   17
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Precio"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   6600
         TabIndex        =   16
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Sub-Total"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid GrDatArticulos 
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   7695
      _ExtentX        =   13573
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
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtorden 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Orden Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   6075
      TabIndex        =   0
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "frm_comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nrosec As String

Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
        Case 0 'Nuevo
            Unload Me
            frm_comp.Show
        Case 1 'Guadar
            activaordenc
            '***************************************************************************
            rsorden.AddNew
            rsorden.Fields("numordenc") = txtorden
            rsorden.Fields("Fecha") = Date
            rsorden.Fields("dir_prov") = txtdatos(2)
            rsorden.Fields("DNI_prov") = txtdatos(0)
           rsorden.Fields("nom_prov") = txtdatos(1)
            'rsfac.Fields ("sub_tot")
            'rsfac.Fields ("IGV")
            rsorden.Fields("total") = txtvalores(2)
            'rsorden.Fields("fec_can") = Date
            'rsfac.Fields ("DNI_per")
            'rsfac.Fields ("nom_per")
            'rsfac.Fields("cod_est") = Datacboest.Text
            'rsfac.Fields("cod_for") = Datacbofor.Text
            
            rsorden.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
           activadorden
            
            Do While Not RsTemporal.EOF
                
                rsdorden.AddNew
                rsdorden.Fields("numordenc") = nrosec
                rsdorden.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                rsdorden.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
               rsdorden.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                rsdorden.Fields("precio") = RsTemporal.Fields("Precio")
               'rsdorden.Fields("Importe") = RsTemporal.Fields("sub_total")
                rsdorden.Update
                RsTemporal.MoveNext
            Loop
            CmdMantenimiento(0).Enabled = True
            CmdMantenimiento(1).Enabled = False
            CmdMantenimiento(2).Enabled = True
            GrDatArticulos.Enabled = False
            
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
                    .Find "cod_pro='" + Trim(dtcproducto.BoundText) + "'"
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
                'txtvalores(0) = FormatCurrency(txtvalores(0), 2)
            End If
            Call GrDatArticulos_Change
    End Select
End Sub



Private Sub dtcproducto_Click(Area As Integer)
rspro.MoveFirst
rspro.Find "des_pro='" + Trim(dtcproducto.Text) + "'"
If rspro.EOF Then
Else

        Txtdescripcion = rspro.Fields("caracteristica")
        TxtPrecio = (rspro.Fields("Pre_ven"))
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
        Txtdescripcion = ""
         TxtPrecio = ""
        txtcantidad = Val(0)
        cmdopciones(0).Enabled = False
        txtcantidad.Enabled = False
     
       
       
        dtcproducto.BoundColumn = "cod_pro"
        dtcproducto.ListField = "des_pro"
        Set dtcproducto.RowSource = rspro
    End If
End Sub

Private Sub Form_Load()
activaordenc
    nrosec = "00-100-000" & (rsorden.RecordCount + 1)
    nrosec = Right(nrosec, 11)
  txtorden = nrosec
    rsorden.Close
    Call ActivaTemporal
    
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
    GrDatArticulos.Columns(2).Caption = "DESCRIPCION"
    GrDatArticulos.Columns(3).Caption = "PRECIO"
    GrDatArticulos.Columns(4).Caption = "CANTIDAD"
    GrDatArticulos.Columns(5).Caption = "SUB TOTAL"
    GrDatArticulos.Columns(1).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(2).Width = 0.5 * GrDatArticulos.Width
    GrDatArticulos.Columns(3).Width = 0.14 * GrDatArticulos.Width
    GrDatArticulos.Columns(4).Width = 0.12 * GrDatArticulos.Width
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
    RsTemporal.Fields(1) = Trim(dtcproducto.BoundText)
    RsTemporal.Fields(2) = Trim(dtcproducto.Text)
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
        'txtvalores(2) = SubTotal
        'txtvalores(2) = Format(txtvalores(2), "##0.00")
     'txtigv = Val(txtvalores(2)) * 0.19
     'txtigv = Format(txtigv, "##0.00")
     
     'txttot.Text = Val(txtvalores(2)) + Val(txtigv.Text)
     'txttot.Text = Format(txttot, "##0.00")
     
    End If
End Sub


Private Sub GrDatArticulos_Change()
If GrDatArticulos.ApproxCount = 0 Then
        'CmdMantenimiento(1).Enabled = False
        'cmdopciones(1).Enabled = False
    Else
        'CmdMantenimiento(1).Enabled = True
        cmdopciones(1).Enabled = True
    End If
    LblContador = "se encuentran profomando " + Str(GrDatArticulos.ApproxCount) + " a"
End Sub


Private Sub Txtcantidad_Change()
 txtsubtotal = (Val(txtcantidad) * (Val(TxtPrecio)))
  txtsubtotal = Format(txtsubtotal, "##0.00")
    cmdopciones(0).Enabled = True
End Sub

Private Sub txtdatos_Change(Index As Integer)
Select Case Index
        Case 0
            If Len(txtdatos(0)) = 8 Then
                activaprov
                rsprov.MoveFirst
                rsprov.Find "DNI_prov ='" + Trim(txtdatos(0)) + "'"
                If Not rsprov.EOF Then
                    txtdatos(1) = rsprov.Fields("nom_prov")
                    txtdatos(2) = rsprov.Fields("dir_prov")
                    'txtdatos(3) = rsprov.Fields("dir_cli")
                    frameCabecera.Enabled = True
                End If
                If Not txtdatos(1) = "" Then
                    txtdatos(1).Enabled = False
                    txtdatos(2).Enabled = False
                    'txtdatos(3).Enabled = False
                Else
                    txtdatos(1).Enabled = True
                    txtdatos(2).Enabled = True
                    'txtdatos(3).Enabled = True
                    txtdatos(1).SetFocus
                    
                    frameCabecera.Enabled = True
                End If
            Else
                txtdatos(1) = ""
                txtdatos(2) = ""
                'txtdatos(3) = ""
                
            End If
    End Select
     If txtdatos(2) <> "" Then
'       Llena el DATALIST (dtlCategoria) con los nombre de categoría
       activalin
        dtllinea.BoundColumn = "cod_lin"
        dtllinea.ListField = "nom_lin"
        Set dtllinea.RowSource = rslin
    End If




End Sub
