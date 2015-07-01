VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmboleta 
   BackColor       =   &H8000000E&
   Caption         =   "Boletas"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   360
      Picture         =   "frmboleta.frx":0000
      TabIndex        =   33
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   960
      Picture         =   "frmboleta.frx":030A
      TabIndex        =   32
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   1560
      Picture         =   "frmboleta.frx":0614
      TabIndex        =   31
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox TxtValores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   2
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame frameCabecera 
      BackColor       =   &H80000009&
      Height          =   2055
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   9375
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
         TabIndex        =   18
         Top             =   960
         Width           =   2880
      End
      Begin VB.TextBox Txtcantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         MaxLength       =   3
         TabIndex        =   17
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox TxtSubTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdOpciones 
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton CmdOpciones 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtcproducto 
         Height          =   315
         Left            =   2520
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2514
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Linea de Productos"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Lista de productos"
         Height          =   195
         Left            =   2640
         TabIndex        =   25
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cant"
         Height          =   195
         Left            =   5520
         TabIndex        =   23
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Precio"
         Height          =   195
         Left            =   6360
         TabIndex        =   22
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Sub-Total"
         Height          =   195
         Left            =   7680
         TabIndex        =   21
         Top             =   120
         Width           =   690
      End
   End
   Begin VB.Frame framePrincipal 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   9375
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   1095
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
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   7320
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "DNI "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Direccion"
         Height          =   195
         Left            =   6600
         TabIndex        =   10
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Distrito"
         Height          =   195
         Left            =   4800
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Nro Boleta"
      Height          =   735
      Left            =   7920
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.TextBox txtbol 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid GrDatArticulos 
      Height          =   975
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   9375
      _ExtentX        =   16536
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
   Begin MSDataListLib.DataCombo Datacbofor 
      Height          =   315
      Left            =   1680
      TabIndex        =   29
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo Datacboest 
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Total"
      Height          =   195
      Left            =   7440
      TabIndex        =   34
      Top             =   4800
      Width           =   360
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
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmboleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nrobol As String

Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
        Case 0 'Nuevo
            Unload Me
            frmboleta.Show
        Case 1 'Guadar
            activabol
            '***************************************************************************
            rsbol.AddNew
            rsbol.Fields("nrobol") = txtbol.Text
            rsbol.Fields("Fec_emi") = Date
            rsbol.Fields("DNI_cli") = txtdatos(0)
            rsbol.Fields("nom_cli") = txtdatos(2)
            
            'rsfac.Fields ("guia_rem")
            'rsfac.Fields ("sub_tot")
            'rsfac.Fields ("IGV")
            rsbol.Fields("importe") = txtvalores(2)
            'rsbol.Fields("fec_can") = Date
            'rsfac.Fields ("DNI_per")
            'rsfac.Fields ("nom_per")
            'rsbol.Fields("cod_est") = Datacboest.Text
            'rsbol.Fields("cod_for") = Datacbofor.Text
            
            rsbol.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
            activadbol
            
            Do While Not RsTemporal.EOF
                
                rsdbol.AddNew
                rsdbol.Fields("nrobol") = txtbol.Text
                rsdbol.Fields("cod_pro") = RsTemporal.Fields("cod_articulo")
                rsdbol.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
                rsdbol.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                'rsdetfac.Fields("precio") = RsTemporal.Fields("Precio")
                rsdbol.Fields("Importe") = RsTemporal.Fields("sub_total")
                rsdbol.Update
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

activabol
nrobol = "00-800-000" & (rsbol.RecordCount + 1)
nrobol = Right(nrobol, 11)
txtbol = nrobol
rsbol.Close

  Call ActivaTemporal

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RsTemporal = Nothing
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
txtsubtotal = (Val(txtcantidad) * (Val(TxtPrecio)))
  txtsubtotal = Format(txtsubtotal, "##0.00")
    cmdopciones(0).Enabled = True
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
    RsTemporal.Fields.Append "cod_articulo", adVarChar, 12, adFldIsNullable
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
        txtvalores(2) = SubTotal
        txtvalores(2) = Format(txtvalores(2), "##0.00")
     
    End If
End Sub
Public Sub SoloNumeros(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8, 46
    Case Is < 48, Is > 57
        KeyAscii = 0
    End Select
End Sub
