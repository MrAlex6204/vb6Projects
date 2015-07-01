VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmguia 
   BackColor       =   &H80000009&
   Caption         =   "Guia Interna Entrega"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   5280
      Picture         =   "frmguia.frx":0000
      TabIndex        =   22
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   5880
      Picture         =   "frmguia.frx":030A
      TabIndex        =   21
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   6480
      Picture         =   "frmguia.frx":0614
      TabIndex        =   20
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtfecha 
      Height          =   285
      Left            =   5400
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtobser 
      Height          =   1095
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "frmguia.frx":091E
      Top             =   3840
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid GrDatArticulos 
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1296
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
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   6855
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdopciones 
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtcantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin MSDataListLib.DataCombo dtcproducto 
         Height          =   315
         Left            =   2520
         TabIndex        =   17
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
         Height          =   1035
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1826
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Linea"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Producto"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6855
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   16
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Encargado"
         Height          =   195
         Left            =   2760
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
      Caption         =   "Nro Guia Interna"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      Begin VB.TextBox txtguiai 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fecha de Entrega"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Observacion"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   900
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
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmguia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nroguiai As String

Private Sub Datacbopro_Click(Area As Integer)
rspro.MoveFirst
des_pro = Datacbopro.Text
rspro.Find "des_pro='" + Trim(Datacbopro.Text) + "'"
If rspro.EOF Then
Else
txtstock.Text = rspro.Fields("stock")
TxtPrecio.Text = rspro.Fields("pre_ven")


End If

End Sub


Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
        Case 0 'Nuevo
            Unload Me
            frmguia.Show
        Case 1 'Guadar
           activaguiai
            '***************************************************************************
            rsguiai.AddNew
            rsguiai.Fields("numnotaguiaint") = txtguiai
            rsguiai.Fields("Fecha") = Date
            rsguiai.Fields("señores") = txtdatos(0)
            rsguiai.Fields("encargado") = txtdatos(1)
            'rsfac.Fields ("guia_rem")
            'rsfac.Fields ("sub_tot")
            'rsfac.Fields ("IGV")
            'rsguiai.Fields("importe") = txtvalores(2)
            rsguiai.Fields("fecha_ent") = txtfecha
            rsguiai.Fields("observacion") = txtobser
            'rsfac.Fields ("nom_per")
            'rsfac.Fields("cod_est") = Datacboest.Text
             'rsfac.Fields("cod_for") = Datacbofor.Text
            
            rsguiai.Update
            '***************************************************************************
              
            RsTemporal.MoveFirst
            '***************************************************************************
            activadguiai
            
            Do While Not RsTemporal.EOF
                
                rsdguiai.AddNew
                rsdguiai.Fields("numnotaguiaint") = nroguiai
                rsdguiai.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                rsdguiai.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
                rsdguiai.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                'rsdetfac.Fields("precio") = RsTemporal.Fields("Precio")
                'rsdetfac.Fields("Importe") = RsTemporal.Fields("sub_total")
                rsdguiai.Update
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
                    .Find "cod_pro='" + Trim(dtcproducto.BoundText) + "'"
                    If .EOF = False Then
                        
                        
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
            
            
            
            GrDatArticulos.Refresh
            Call GrDatArticulos_Change
            
            Case 1 'ELIMINAR ITEM
            If RsTemporal.RecordCount > 0 Then
                If RsTemporal.BOF Or RsTemporal.EOF Then MsgBox "Debe seleccionar de la grilla el Titulo a eliminar", vbCritical, "ACONT": Exit Sub
                RsTemporal.Delete
                RsTemporal.MoveNext
                
                If RsTemporal.RecordCount = 0 Then
                   
                    RsTemporal.Close
                    ActivaTemporal
                End If
            Else
                cmdopciones(1).Enabled = False
                
               
                
            End If
            Call GrDatArticulos_Change
    End Select
End Sub

Private Sub dtcproducto_Change()
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
        'Txtdescripcion = ""
         'TxtPrecio = ""
        txtcantidad = Val(0)
        'cmdopciones(0).Enabled = False
        txtcantidad.Enabled = False
     
       
       
        dtcproducto.BoundColumn = "cod_pro"
        dtcproducto.ListField = "des_pro"
        Set dtcproducto.RowSource = rspro
    End If
End Sub

Private Sub Form_Load()
activaguiai

nroguiai = "100-200-00" & (rsguiai.RecordCount + 1)
    nroguiai = Right(nroguiai, 11)
   txtguiai = nroguiai
    rsguiai.Close

Call ActivaTemporal

txtfecha.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub imgCalendario_Click(Index As Integer)
Dim vFecha As Variant

  imgCalendario(0).Refresh
  If IsDate(mebFecha.Text) Then
    vFecha = fvGetDate(mebFecha.Text)
  Else
    vFecha = fvGetDate(Date)
  End If
  If vFecha <> False Then
    mebFecha.Text = Format(vFecha, "dd/mm/yyyy")
  End If
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
Sub ActivaTemporal()
    'CREANDO RECORDSET TEMPORAL****************
    Set RsTemporal = New ADODB.Recordset
    RsTemporal.CursorType = adOpenStatic
    RsTemporal.Fields.Append "IdEnProf", adVarChar, 10, adFldIsNullable
    RsTemporal.Fields.Append "cod_art", adVarChar, 12, adFldIsNullable
    RsTemporal.Fields.Append "Nom_Art", adVarChar, 250, adFldIsNullable
    RsTemporal.Fields.Append "cantidad", adInteger, adFldIsNullable
    RsTemporal.Open
    '*****************************************************
     Set GrDatArticulos.DataSource = RsTemporal
     GrDatArticulos.Columns(0).Visible = False
    GrDatArticulos.Columns(1).Visible = False
    
    GrDatArticulos.Columns(1).Caption = "ITEMS"
    GrDatArticulos.Columns(2).Caption = "DESCRIPCION"
    GrDatArticulos.Columns(3).Caption = "CANTIDAD"
    GrDatArticulos.Columns(1).Width = 0.1 * GrDatArticulos.Width
    GrDatArticulos.Columns(2).Width = 0.5 * GrDatArticulos.Width
     GrDatArticulos.Columns(3).Width = 0.17 * GrDatArticulos.Width
    GrDatArticulos.Columns(3).Alignment = dbgRight

End Sub
Sub Graba_Temporal()
    RsTemporal.AddNew
    RsTemporal.Fields(0) = Trim(txtnumero)
    RsTemporal.Fields(1) = Trim(dtcproducto.BoundText)
    RsTemporal.Fields(2) = Trim(dtcproducto.Text)
    'RsTemporal.Fields(3) = Trim(TxtPrecio)
    RsTemporal.Fields(3) = Trim(txtcantidad)
    'RsTemporal.Fields(5) = Trim(txtsubtotal)
    RsTemporal.Update
End Sub
