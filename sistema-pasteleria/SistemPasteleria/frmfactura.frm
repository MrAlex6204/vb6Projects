VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmfactura 
   BackColor       =   &H8000000E&
   Caption         =   "Factura"
   ClientHeight    =   6045
   ClientLeft      =   1500
   ClientTop       =   2895
   ClientWidth     =   9585
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9585
   Begin VB.CommandButton Command1 
      Caption         =   "::"
      Height          =   255
      Left            =   9240
      TabIndex        =   45
      Top             =   840
      Width           =   255
   End
   Begin VB.Frame frameDetalle 
      BackColor       =   &H8000000E&
      Height          =   2415
      Left            =   240
      TabIndex        =   20
      Top             =   3600
      Width           =   9255
      Begin MSDataGridLib.DataGrid GrDatArticulos 
         Height          =   975
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
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
      Begin VB.TextBox TxtValores 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   2
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1200
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo Datacbofor 
         Height          =   315
         Left            =   1560
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Datacboest 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtigv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   7560
         TabIndex        =   25
         Top             =   1560
         Width           =   1090
      End
      Begin VB.TextBox txttot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   7560
         TabIndex        =   24
         Top             =   1920
         Width           =   1090
      End
      Begin VB.CommandButton CmdMantenimiento 
         BackColor       =   &H80000009&
         Caption         =   "PREVIO"
         Enabled         =   0   'False
         Height          =   615
         Index           =   2
         Left            =   1440
         Picture         =   "frmfactura.frx":0000
         TabIndex        =   23
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton CmdMantenimiento 
         BackColor       =   &H80000009&
         Caption         =   "GUARDAR"
         Enabled         =   0   'False
         Height          =   615
         Index           =   1
         Left            =   840
         Picture         =   "frmfactura.frx":030A
         TabIndex        =   22
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton CmdMantenimiento 
         BackColor       =   &H80000009&
         Caption         =   "NUEVO"
         Enabled         =   0   'False
         Height          =   615
         Index           =   0
         Left            =   240
         Picture         =   "frmfactura.frx":0614
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   1560
         TabIndex        =   32
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Estado"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Importe"
         Height          =   195
         Left            =   6960
         TabIndex        =   28
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "IGV"
         Height          =   195
         Left            =   7200
         TabIndex        =   27
         Top             =   1560
         Width           =   270
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Total"
         Height          =   195
         Left            =   7080
         TabIndex        =   26
         Top             =   1920
         Width           =   360
      End
   End
   Begin VB.Frame frameCabecera 
      BackColor       =   &H80000009&
      Height          =   2055
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   9255
      Begin VB.TextBox txtstock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5280
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CmdOpciones 
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   40
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdOpciones 
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   39
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxtSubTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   1095
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Txtcantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   36
         Top             =   360
         Width           =   720
      End
      Begin VB.TextBox Txtdescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   975
         Left            =   2520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   960
         Width           =   2880
      End
      Begin MSDataListLib.DataCombo dtcproducto 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
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
         Height          =   1425
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2514
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Stock"
         Height          =   195
         Left            =   5400
         TabIndex        =   50
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Sub-Total"
         Height          =   195
         Left            =   8160
         TabIndex        =   19
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Precio"
         Height          =   195
         Left            =   7080
         TabIndex        =   18
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cant"
         Height          =   195
         Left            =   6120
         TabIndex        =   17
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Lista de productos"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Linea de Productos"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame framePrincipal 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   9255
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   7320
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5400
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Distrito"
         Height          =   195
         Left            =   4800
         TabIndex        =   9
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Direccion"
         Height          =   195
         Left            =   6600
         TabIndex        =   4
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "DNI "
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Nro Factura"
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.TextBox TxtNumeroProforma 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Hora"
      Height          =   195
      Left            =   3120
      TabIndex        =   52
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Fecha"
      Height          =   195
      Left            =   3120
      TabIndex        =   51
      Top             =   600
      Width           =   450
   End
   Begin VB.Label lblguia 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   6600
      TabIndex        =   48
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   47
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guia de Remision"
      Height          =   195
      Left            =   7920
      TabIndex        =   46
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label lblven 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   3120
      TabIndex        =   43
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblfecha 
      BackColor       =   &H80000009&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   42
      Top             =   600
      Width           =   3135
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
      Left            =   2400
      TabIndex        =   33
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmfactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nrosec As String

Private Sub cmdAyuda_Click()
activacli
    dtlcliente.BoundColumn = "DNI_cli"
    dtlcliente.ListField = "nom_cli"
    Set dtlcliente.RowSource = rscli
    dtlcliente.Visible = True
    cmdAyuda.Enabled = False

End Sub




Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
        Case 0 'Nuevo
            Unload Me
            frmfactura.Show
        Case 1 'Guadar
            activafac
            '***************************************************************************
            rsfac.AddNew
            rsfac.Fields("nrofac") = TxtNumeroProforma
            rsfac.Fields("Fec_emi") = Date
            rsfac.Fields("nom_cli") = txtdatos(2)
            rsfac.Fields("Dni_cli") = txtdatos(0)
            'rsfac.Fields ("guia_rem")
            rsfac.Fields("sub_tot") = txtsubtotal.Text
            rsfac.Fields("IGV") = txtigv.Text
            rsfac.Fields("total") = txttot.Text
            rsfac.Fields("fec_can") = Date
            'rsfac.Fields ("DNI_per")
            'rsfac.Fields ("nom_per")
            rsfac.Fields("cod_est") = Datacboest.Text
            rsfac.Fields("cod_for") = Datacbofor.Text
            rsfac.Fields("Emitido") = "N"
            rsfac.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
            activadetfac
            
            Do While Not RsTemporal.EOF
                
                rsdetfac.AddNew
                rsdetfac.Fields("nrofac") = nrosec
                rsdetfac.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                rsdetfac.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
                rsdetfac.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                'rsdetfac.Fields("precio") = RsTemporal.Fields("Precio")
                rsdetfac.Fields("Importe") = RsTemporal.Fields("sub_total")
                rsdetfac.Update
                RsTemporal.MoveNext
            Loop
            CmdMantenimiento(0).Enabled = True
            CmdMantenimiento(1).Enabled = False
            CmdMantenimiento(2).Enabled = True
            GrDatArticulos.Enabled = False
            framePrincipal.Enabled = False
            frameCabecera.Enabled = False
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
                    .Find "cod_art='" + Trim(dtcproducto.BoundText) + "'"
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
            frameDetalle.Enabled = True
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

Private Sub Command1_Click()
frmguiaremision.Show
End Sub

Private Sub datacboest_Click(Area As Integer)
rsest.MoveFirst
nom_est = Datacboest.Text
rsest.Find "nom_est='" + Trim(Datacboest.Text) + "'"
If rsest.EOF Then

Else

End If

End Sub

Private Sub dtcproducto_Change()
rspro.MoveFirst
rspro.Find "des_pro='" + Trim(dtcproducto.Text) + "'"
If rspro.EOF Then
Else

        Txtdescripcion = rspro.Fields("caracteristica")
        TxtPrecio = (rspro.Fields("Pre_ven"))
        txtstock = (rspro.Fields("stock"))
        txtstock.Enabled = False
        txtcantidad.Enabled = True
        txtcantidad = Val(0)
        txtcantidad.SetFocus

End If
End Sub



Private Sub dtlcliente_Click()
 rscli.MoveFirst
   rscli.Find "DNI_cli ='" + Trim(dtlcliente.BoundText) + "'"
    If Not rscli.EOF Then
        txtdatos(0) = rscli.Fields("DNI_cli")
        txtdatos(1) = rscli.Fields("cod_dis")
        txtdatos(2) = rscli.Fields("nom_cli")
        txtdatos(3) = rscli.Fields("dir_cli")
    End If
    rscli.Close
    dtlcliente.Enabled = True
    dtlcliente.Visible = False
    cmdAyuda.Enabled = True
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
        Datacboest.Enabled = False
       Datacbofor.Enabled = False
       
        dtcproducto.BoundColumn = "cod_pro"
        dtcproducto.ListField = "des_pro"
        Set dtcproducto.RowSource = rspro
    End If
End Sub

Private Sub Form_Activate()
frmfactura.Top = 0
    frmfactura.Left = 0
End Sub
Private Sub Form_Load()


activafac
    nrosec = "00-100-000" & (rsfac.RecordCount + 1)
    nrosec = Right(nrosec, 11)
    TxtNumeroProforma = nrosec
    rsfac.Close


lblfecha = Format(Date, "long date")
lblhora = Format(Time, "long time")

    Call ActivaTemporal

 activaest
 Datacboest.ListField = "nom_est"
 Set Datacboest.RowSource = rsest
 
 activafor
 Datacbofor.ListField = "nom_for"
 Set Datacbofor.RowSource = rsfor
 

activapro
dtcproducto.ListField = "des_pro"
Set dtcproducto.RowSource = rspro

frmfactura.lblven = frmreg.txtDNI(1)
    

frmfactura.lblguia = frmguiaremision.txtguiar

    
    
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



Private Sub Form_Unload(Cancel As Integer)
 Set RsTemporal = Nothing
End Sub






Private Sub Txtcantidad_Change()
  txtsubtotal = (Val(txtcantidad) * (Val(TxtPrecio)))
  txtsubtotal = Format(txtsubtotal, "##0.00")
  txtstock = (Val(txtstock) - (Val(txtcantidad)))
  If Val(txtstock) < Val(txtcantidad) Then
  MsgBox "CANTIDAD DE STOCK NO PERMITIDA", vbCritical, "SISTEMA DE SEGURIDAD"
    txtstock = rspro.Fields("stock")
    txtcantidad.Text = ""
    txtcantidad.SetFocus
    End If
    If txtcantidad = "" Then
    txtstock.Text = rspro.Fields("stock")
    End If
    cmdopciones(0).Enabled = True
    Datacboest.Enabled = True
   Datacbofor.Enabled = True
End Sub

Private Sub txtdatos_Change(Index As Integer)

Select Case Index
        Case 0
            If Len(txtdatos(0)) = 8 Then
                activacli
                rscli.MoveFirst
                rscli.Find "DNI_cli ='" + Trim(txtdatos(0)) + "'"
                If Not rscli.EOF Then
                    txtdatos(1) = rscli.Fields("cod_dis")
                    txtdatos(2) = rscli.Fields("nom_cli")
                    txtdatos(3) = rscli.Fields("dir_cli")
                    frameCabecera.Enabled = True
                End If
                If Not txtdatos(1) = "" Then
                    txtdatos(1).Enabled = False
                    txtdatos(2).Enabled = False
                    txtdatos(3).Enabled = False
                Else
                    txtdatos(1).Enabled = True
                    txtdatos(2).Enabled = True
                    txtdatos(3).Enabled = True
                    txtdatos(1).SetFocus
                    
                    frameCabecera.Enabled = True
                End If
            Else
                txtdatos(1) = ""
                txtdatos(2) = ""
                txtdatos(3) = ""
                
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
        txtvalores(2) = SubTotal
        txtvalores(2) = Format(txtvalores(2), "##0.00")
     txtigv = Val(txtvalores(2)) * 0.19
     txtigv = Format(txtigv, "##0.00")
     
     txttot.Text = Val(txtvalores(2)) + Val(txtigv.Text)
     txttot.Text = Format(txttot, "##0.00")
     
    End If
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
        Case 0
            Call SoloNumeros(KeyAscii)
    End Select
    
End Sub

Public Sub SoloNumeros(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8, 46
    Case Is < 48, Is > 57
        KeyAscii = 0
    End Select
End Sub




