VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpago 
   BackColor       =   &H80000009&
   Caption         =   "Pagos a Cuenta"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   1440
      Width           =   6015
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   25
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Cliente"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "DNI"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   285
      End
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "NUEVO"
      Height          =   615
      Index           =   0
      Left            =   4320
      Picture         =   "frmpago.frx":0000
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "GUARDAR"
      Height          =   615
      Index           =   1
      Left            =   4920
      Picture         =   "frmpago.frx":030A
      TabIndex        =   19
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton CmdMantenimiento 
      BackColor       =   &H80000009&
      Caption         =   "PREVIO"
      Height          =   615
      Index           =   2
      Left            =   5520
      Picture         =   "frmpago.frx":0614
      TabIndex        =   18
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtsaldo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "---"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdAyuda 
      Caption         =   "?"
      Height          =   255
      Left            =   5745
      TabIndex        =   13
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtNro 
      Alignment       =   2  'Center
      Height          =   264
      Left            =   4320
      MaxLength       =   11
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtmonto 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5040
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtfecha 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtnump 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5040
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtcue 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin MSDataListLib.DataList dtlProforma 
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Nro Factura"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "PANADERIA PASTELERIA ALISSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   4140
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   5160
      TabIndex        =   7
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   510
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   750
   End
End
Attribute VB_Name = "frmpago"
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
        Case 0 'Nuevo
            Unload Me
            frmpago.Show
        Case 1 'Guadar
            activapago
            '***************************************************************************
            rspago.AddNew
            rspago.Fields("numpago") = txtnump
            rspago.Fields("Fec_emi") = Date
            rspago.Fields("DNI_cli") = txtdatos(0)
            rspago.Fields("nom_cli") = txtdatos(1)
            rspago.Fields("total") = txttotal
            rspago.Fields("acuenta") = txtcue
            rspago.Fields("saldo") = txtsaldo
            rspago.Fields("monto") = txtmonto
            rspago.Fields("fecha_apagar") = txtfecha
            'rsfac.Fields ("DNI_per")
            'rsfac.Fields ("nom_per")
            
            rspago.Update
            '***************************************************************************
              
            'RsTemporal.MoveFirst
            '***************************************************************************
            'activadetfac
            
            'Do While Not RsTemporal.EOF
                
                'rsdetfac.AddNew
                'rsdetfac.Fields("nrofac") = nrosec
                'rsdetfac.Fields("cod_pro") = RsTemporal.Fields("cod_art")
                'rsdetfac.Fields("nom_pro") = RsTemporal.Fields("Nom_Art")
                'rsdetfac.Fields("cantidad") = RsTemporal.Fields("Cantidad")
                'rsdetfac.Fields("precio") = RsTemporal.Fields("Precio")
                'rsdetfac.Fields("Importe") = RsTemporal.Fields("sub_total")
                'rsdetfac.Update
                'RsTemporal.MoveNext
            'Loop
            'CmdMantenimiento(0).Enabled = True
            'CmdMantenimiento(1).Enabled = False
            'CmdMantenimiento(2).Enabled = True
            'GrDatArticulos.Enabled = False
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

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Dim a, b, c As Double

a = Val(txttotal.Text)
b = Val(txtcue.Text)
c = a - b

txtsaldo.Text = Format(c, "##0.00")

txtmonto.Text = txtsaldo.Text
txtmonto.Text = Format(txtmonto, "##0.00")
    
If txttotal.Text < txtcue.Text Then
MsgBox "UD A INGRESADO MAS DE LO QUE RESTA", vbCritical, "SISTEMA DE SEGURIDAD"
txtcue.Text = ""
txtsaldo.Text = ""
txtmonto.Text = ""
txtcue.SetFocus
End If


End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub dtlProforma_Click()
TxtNro = dtlProforma.Text
    rsfac.Close
    dtlProforma.Visible = False
    cmdAyuda.Enabled = True
End Sub

Private Sub Form_Load()
activapago
nrosec = "100-200-00" & (rspago.RecordCount + 1)
nrosec = Right(nrosec, 11)
txtnump.Text = nrosec
rspago.Close



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
            
            
        Else
            Nro = rsfac.Fields("Emitido")
            If Nro = "N" Then
                    txtdatos(0) = rsfac.Fields("DNI_cli")
                    txtdatos(1) = rsfac.Fields("nom_cli")
                    txttotal.Text = rsfac.Fields("total")
                    txttotal = Format(txttotal, "##0.00")
                    txtcue.SetFocus
                     ' **********************************
                    Else
                    MsgBox "El número de proforma ingresado no existe o ya se ha Emitido de la base de datos.", vbCritical, "Sistema de facturación"
                    TxtNro.SelStart = 0
                    TxtNro.SelLength = Len(TxtNro)
                    
                    
                    
                    
            End If
        End If
    Else
        
    End If
End Sub
