VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmguiaremision 
   BackColor       =   &H80000009&
   Caption         =   "Guia de Remision"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   46
      Top             =   5280
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   45
      Top             =   4920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   44
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdmantenimiento 
      Caption         =   "NUEVO"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   27
      Top             =   5760
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid GrDatFactura 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1508
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
   Begin VB.TextBox txttot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   22
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtigv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8040
      TabIndex        =   21
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox TxtValores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H00000040&
      Height          =   285
      Index           =   2
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdmantenimiento 
      Caption         =   "PREVIO"
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdmantenimiento 
      Caption         =   "GUARDAR"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   5760
      Width           =   615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   9375
      Begin VB.TextBox txtplaca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   40
         Top             =   960
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo datacboveh 
         Height          =   315
         Left            =   7560
         TabIndex        =   39
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   7200
         TabIndex        =   38
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   37
         Top             =   120
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo datacbodisl 
         Height          =   315
         Left            =   4320
         TabIndex        =   35
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo datacbodiss 
         Height          =   315
         Left            =   4320
         TabIndex        =   34
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtllega 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtsal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   480
         MaxLength       =   8
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "T.Vehiculo"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6720
         TabIndex        =   36
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Distrito"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3720
         TabIndex        =   31
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Distrito"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "P.Llegada"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "P.Salida"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Marca 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RUC"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6720
         TabIndex        =   19
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Placa"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6720
         TabIndex        =   18
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Transportista"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   915
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "DNI"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9375
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "DNI"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Cliente"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   1680
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Nro Guia Remision"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   7800
      TabIndex        =   4
      Top             =   240
      Width           =   1695
      Begin VB.TextBox txtguiar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame framePrincipal 
      BackColor       =   &H80000009&
      Caption         =   "Nro Factura"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.TextBox TxtNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   264
         Left            =   90
         MaxLength       =   11
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "?"
         Height          =   255
         Left            =   1515
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin MSDataListLib.DataList dtlProforma 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Fecha Entrega"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Motivo Traslado"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   4920
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Persona  Despacha"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   1410
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Total"
      Height          =   255
      Left            =   7320
      TabIndex        =   25
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "IGV"
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Importe"
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   4560
      Width           =   615
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
      TabIndex        =   15
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmguiaremision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim nrosec As String
Private Sub cmdAyuda_Click()
Set rsfac = New ADODB.Recordset
rsfac.Open "SELECT * FROM facturas WHERE Emitido='" + Trim("N") + "'", cn, adOpenStatic, adLockOptimistic
    dtlProforma.BoundColumn = "nrofac"
    dtlProforma.ListField = "nrofac"
    Set dtlProforma.RowSource = rsfac
    dtlProforma.Visible = True
    cmdAyuda.Enabled = False
End Sub

Private Sub CmdMantenimiento_Click(Index As Integer)
 Select Case Index
    Case 0
        activaguia
        ' ENCABEZADO guia remision
        
        rsguia.AddNew
        rsguia.Fields("numguiarem") = Trim(txtguiar)
        rsguia.Fields("Fecha_emi") = Date
        rsguia.Fields("subtotal") = (txtvalores(2))
        'rsguia.Fields("Anulado") = "N"
        'rsguia.Fields("Desc_Bol") = 0
        rsguia.Fields("DNI_cli") = txtdatos(1)
        'rsdguia.Fields("nom_cli") = TxtDatos(2)
        rsguia.Update
        cn.Execute "Update facturas Set Emitido='S' Where nrofac='" & Trim(TxtNro) & "'"
         '***************************************************************************
         
         
         
         
         
       ' DETALLE DE LA guia remision
       activadguia
        RsTemporal.MoveFirst
        Do While Not RsTemporal.EOF
            rsdguia.AddNew
            rsdguia.Fields("numguiarem") = Trim(txtguiar)
            rsdguia.Fields("cod_pro") = RsTemporal.Fields("IdArticulo")
            rsdguia.Fields("Cantidad") = RsTemporal.Fields("Cantidad")
            'rsdguia.Fields("Prec_det") = RsTemporal.Fields("Precio")
            rsdguia.Fields("Importe") = RsTemporal.Fields("sub_total")
            rsdguia.Update
            RsTemporal.MoveNext
        Loop
        CmdMantenimiento(0).Enabled = False
        CmdMantenimiento(1).Enabled = True
        framePrincipal.Enabled = False
    
    Case 1
        
        If DataEnvironment1.rscmdguiaremision.State = adStateOpen Then
           DataEnvironment1.rscmdguiaremision.Close
        End If
        DataEnvironment1.cmdguiaremision (Trim(txtguiar))
        
        Set dtaguiar.DataSource = DataEnvironment1.rscmdguiaremision
        dtaguiar.Caption = "Guia de Remision"
        dtaguiar.Show
    
    End Select



End Sub


Private Sub datacbodisl_Click(Area As Integer)
rsdis.MoveFirst
nom_dis = datacbodisl.Text
rsdis.Find "nom_dis='" + Trim(datacbodisl.Text) + "'"
If rsdis.EOF Then
Else
End If

End Sub

Private Sub datacbodiss_Click(Area As Integer)
rsdis.MoveFirst
nom_dis = datacbodiss.Text
rsdis.Find "nom_dis='" + Trim(datacbodiss.Text) + "'"
If rsdis.EOF Then
Else
End If

End Sub

Private Sub datacboveh_Click(Area As Integer)
rsve.MoveFirst
tipo = datacboveh.Text
rsve.Find "tipo='" + Trim(datacboveh.Text) + "'"
If rsve.EOF Then
Else
txtplaca.Text = rsve.Fields("num_veh")
End If
End Sub

Private Sub dtlProforma_Click()
TxtNro = dtlProforma.Text
    rsfac.Close
    dtlProforma.Visible = False
    cmdAyuda.Enabled = True
End Sub



Private Sub Form_Load()
activaguia
    nrosec = "00-500-000" & (rsguia.RecordCount + 1)
    nrosec = Right(nrosec, 11)
    txtguiar = nrosec
    rsguia.Close
activadis
datacbodiss.ListField = "nom_dis"
Set datacbodiss.RowSource = rsdis

activadis
datacbodisl.ListField = "nom_dis"
Set datacbodisl.RowSource = rsdis

activave
datacboveh.ListField = "tipo"
Set datacboveh.RowSource = rsve
End Sub

Private Sub txtdato_Change(Index As Integer)
Select Case Index
        Case 0
            If Len(txtdato(0)) = 8 Then
                activatrans
                rstrans.MoveFirst
                rstrans.Find "DNI_trans ='" + Trim(txtdato(0)) + "'"
                If Not rstrans.EOF Then
                    txtdato(1) = rstrans.Fields("nom_trans")
                    txtdato(2) = rstrans.Fields("RUC_trans")
                    
                    'frameCabecera.Enabled = True
                End If
                If Not txtdatos(1) = "" Then
                    txtdato(1).Enabled = False
                    txtdato(2).Enabled = False
                                    Else
                    txtdato(1).Enabled = True
                    txtdato(2).Enabled = True
                    txtdato(1).SetFocus
                    
                    frameCabecera.Enabled = True
                End If
            Else
                txtdato(1) = ""
                txtdato(2) = ""
    End If
    End Select
    txtsal.SetFocus
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
     
End Sub

Private Sub TxtNro_Change()
If Len(TxtNro) = 11 Then
        Dim rsfac As New Recordset
        Set rsfac = New ADODB.Recordset
        rsfac.Open "SELECT * FROM facturas WHERE nrofac='" + Trim(TxtNro) + "'", cn, adOpenStatic, adLockOptimistic
        If rsfac.BOF Then
            MsgBox "El número de proforma ingresado no existe o ya se ha Emitido de la base de datos.", vbCritical, "Sistema de facturación"
            Set RsTemporal = Nothing
            TxtNro.SelStart = 0
            TxtNro.SelLength = Len(TxtNro)
            CmdMantenimiento(0).Enabled = False
            txtdatos(1) = ""
            txtdatos(2) = ""
            
        Else
            Nro = rsfac.Fields("Emitido")
            If Nro = "N" Then
                    txtdatos(1) = rsfac.Fields("DNI_cli")
                    txtdatos(2) = rsfac.Fields("nom_cli")
                    Call ActivaTemporal
                    Call Calculos
                    ' **********************************
                    Dim rscli As New Recordset
                    rscli.Open "SELECT * FROM Clientes WHERE DNI_cli ='" + Trim(txtdatos(1)) + "' AND nom_cli = '" + Trim(txtdatos(2)) + "'", cn, adOpenStatic, adLockOptimistic
                    If rscli.BOF Then MsgBox "CLIENTE NO VALIDO": Exit Sub
                    Cod_Cliente = rscli.Fields("DNI_cli")
                    ' **********************************
                    CmdMantenimiento(0).Enabled = True
            Else
                    MsgBox "El número de proforma ingresado no existe o ya se ha Emitido de la base de datos.", vbCritical, "Sistema de facturación"
                    Set RsTemporal = Nothing
                    TxtNro.SelStart = 0
                    TxtNro.SelLength = Len(TxtNro)
                    CmdMantenimiento(0).Enabled = False
                    txtdatos(1) = ""
                    txtdatos(2) = ""
            End If
        End If
    Else
        CmdMantenimiento(1).Enabled = False
        Set GrDatFactura.DataSource = Nothing
        Set RsTemporal = Nothing
        GrDatFactura.Refresh
    End If
End Sub

Sub ActivaTemporal()
    Set RsTemporal = New Recordset
    RsTemporal.CursorType = adOpenStatic
    RsTemporal.Fields.Append "IdEnProf", adVarChar, 10, adFldIsNullable
    RsTemporal.Fields.Append "IdArticulo", adVarChar, 12, adFldIsNullable
    RsTemporal.Fields.Append "Nom_Art", adVarChar, 250, adFldIsNullable
    RsTemporal.Fields.Append "Precio", adDouble, 10.2, adFldIsNullable
    RsTemporal.Fields.Append "cantidad", adInteger, adFldIsNullable
    RsTemporal.Fields.Append "Sub_total", adDouble, 10.2, adFldIsNullable
    RsTemporal.Open
    Dim rspro  As New Recordset
    rspro.Open "SELECT * FROM productos", cn, adOpenStatic, adLockBatchOptimistic
    
    'detalle de guia de factura
    
    
    Dim rsdfac As Recordset
    Set rsdfac = New ADODB.Recordset
    rsdfac.Open "Select * from Detallefacturas Where nrofac='" & Trim(TxtNro) & "'", cn, adOpenStatic, adLockBatchOptimistic
    rsdfac.MoveFirst
    Do
            RsTemporal.AddNew
            
            'RsTemporal("IdEnProf") = rsdfac("nrofac")
            
            RsTemporal("idarticulo") = rsdfac("cod_pro")
            ' RECUPERA EL NOMBRE DE ARTICULO ****************************
            rspro.MoveFirst
            rspro.Find "cod_pro ='" + Trim(rsdfac("nrofac")) + "'"
            If Not rspro.EOF Then RsTemporal("Nom_Art") = rspro.Fields("nom_pro")
            ' *********************************************************
            RsTemporal("Nom_Art") = rsdfac.Fields("nom_pro")
            'RsTemporal("precio") = rsdfac.Fields("precio")
            RsTemporal("Cantidad") = rsdfac("cantidad")
            RsTemporal("sub_total") = rsdfac("Importe")
            RsTemporal.Update
            rsdfac.MoveNext
       Loop While Not rsdfac.EOF
    Set GrDatFactura.DataSource = RsTemporal
    
    'DATA GRID
    
    Set GrDatFactura.DataSource = RsTemporal
     GrDatFactura.Columns(0).Visible = False
    GrDatFactura.Columns(1).Visible = False
    
    GrDatFactura.Columns(1).Caption = "ITEMS"
    GrDatFactura.Columns(2).Caption = "DESCRIPCION"
    GrDatFactura.Columns(3).Caption = "PRECIO"
    GrDatFactura.Columns(4).Caption = "CANTIDAD"
    GrDatFactura.Columns(5).Caption = "SUB TOTAL"
   GrDatFactura.Columns(1).Width = 0.1 * GrDatFactura.Width
   GrDatFactura.Columns(2).Width = 0.5 * GrDatFactura.Width
  GrDatFactura.Columns(3).Width = 0.14 * GrDatFactura.Width
 GrDatFactura.Columns(4).Width = 0.12 * GrDatFactura.Width
  GrDatFactura.Columns(5).Width = 0.2 * GrDatFactura.Width
GrDatFactura.Columns(3).NumberFormat = "##,##0.00"
GrDatFactura.Columns(5).NumberFormat = "##,##0.00"
GrDatFactura.Columns(3).Alignment = dbgRight
GrDatFactura.Columns(4).Alignment = dbgRight
GrDatFactura.Columns(5).Alignment = dbgRight

    
    
    
    
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

