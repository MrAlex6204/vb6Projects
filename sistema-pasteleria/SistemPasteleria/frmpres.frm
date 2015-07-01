VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpresentacion 
   BackColor       =   &H8000000C&
   Caption         =   "Presentacion"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   4200
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7320
      Top             =   1560
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16776960
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":1FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmpres.frx":260C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Negrita"
            Object.ToolTipText     =   "Negrita"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cursiva"
            Object.ToolTipText     =   "Cursiva"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Subrayado"
            Object.ToolTipText     =   "Subrayado"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alinear a la izquierda"
            Object.ToolTipText     =   "Alinear a la izquierda"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Centrar"
            Object.ToolTipText     =   "Centrar"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alinear a la derecha"
            Object.ToolTipText     =   "Alinear a la derecha"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "PANADERIA PASTELERIA ""ALISON"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1920
      TabIndex        =   2
      Top             =   3960
      Width           =   9405
   End
   Begin VB.Image Image01 
      Height          =   1005
      Left            =   480
      Picture         =   "frmpres.frx":2A5E
      Top             =   960
      Width           =   1800
   End
   Begin VB.Image Image02 
      Height          =   1005
      Left            =   9240
      Picture         =   "frmpres.frx":5162
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label lbltitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4320
      TabIndex        =   0
      Top             =   840
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   -600
      Picture         =   "frmpres.frx":7866
      Top             =   -1560
      Width           =   1335
   End
   Begin VB.Menu procesos 
      Caption         =   "PROCESOS"
      Begin VB.Menu factura 
         Caption         =   "Facturas"
      End
      Begin VB.Menu boletas1 
         Caption         =   "Boletas"
      End
      Begin VB.Menu notapedidointerno 
         Caption         =   "Nota pedido Interno"
      End
      Begin VB.Menu pagocuenta 
         Caption         =   "Pago a Cuenta"
      End
      Begin VB.Menu ingresodedocumento 
         Caption         =   "Ingreso de Documento"
      End
      Begin VB.Menu OrdenCompra 
         Caption         =   "Orden Compra"
      End
      Begin VB.Menu notaguiainterna 
         Caption         =   "Nota Guia Interna"
      End
   End
   Begin VB.Menu tablas 
      Caption         =   "TABLAS"
      Begin VB.Menu clientes1 
         Caption         =   "Clientes"
      End
      Begin VB.Menu distritos 
         Caption         =   "Distritos"
      End
      Begin VB.Menu proveedores1 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu Productos1 
         Caption         =   "Productos"
      End
      Begin VB.Menu cargos 
         Caption         =   "Cargos"
      End
      Begin VB.Menu marcas 
         Caption         =   "Marcas"
      End
   End
   Begin VB.Menu consultas 
      Caption         =   "CONSULTAS"
      Begin VB.Menu vehiculos 
         Caption         =   "Vehiculos"
      End
      Begin VB.Menu proveedores 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu productos 
         Caption         =   "Productos"
      End
      Begin VB.Menu OrdendePedido 
         Caption         =   "Orden de Pedido"
      End
      Begin VB.Menu marca 
         Caption         =   "Marcas"
      End
      Begin VB.Menu pagosacuenta 
         Caption         =   "Pagos a Cuenta"
      End
      Begin VB.Menu distrito 
         Caption         =   "Distrito"
      End
      Begin VB.Menu cliente 
         Caption         =   "Cliente"
      End
   End
   Begin VB.Menu reportes 
      Caption         =   "REPORTES"
      Begin VB.Menu personales 
         Caption         =   "Personal"
      End
      Begin VB.Menu forma 
         Caption         =   "Forma de Pago"
      End
      Begin VB.Menu lineasr 
         Caption         =   "Lineas"
      End
      Begin VB.Menu EntregaInternasporFechas 
         Caption         =   "Entrega Internas por Fechas"
      End
      Begin VB.Menu ProductosenLinea 
         Caption         =   "Productos en Linea"
      End
      Begin VB.Menu GeneraldeProductos 
         Caption         =   "General de Productos"
      End
      Begin VB.Menu facturasporfechas1 
         Caption         =   "Facturas por Fechas"
      End
      Begin VB.Menu BoletasporFechas 
         Caption         =   "Boletas por Fechas"
      End
      Begin VB.Menu PagosaCuentaFecha 
         Caption         =   "Pagos a Cuenta Fecha"
      End
      Begin VB.Menu clientes 
         Caption         =   "Clientes por Distrito"
      End
   End
   Begin VB.Menu graficos 
      Caption         =   "GRAFICOS"
      Begin VB.Menu ProductomasVendido 
         Caption         =   "Producto mas Vendido"
      End
      Begin VB.Menu ventasporVendedor 
         Caption         =   "Ventas por Vendedor"
      End
      Begin VB.Menu clientes2 
         Caption         =   "Clientes"
      End
      Begin VB.Menu facturasporfechas 
         Caption         =   "Facturas a la Fecha"
      End
      Begin VB.Menu ventas1 
         Caption         =   "Ventas"
      End
   End
   Begin VB.Menu mantenimiento 
      Caption         =   "MANTENIMIENTO"
      Begin VB.Menu backup 
         Caption         =   "Backup"
      End
      Begin VB.Menu restore 
         Caption         =   "Restore"
      End
      Begin VB.Menu usuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu personal 
         Caption         =   "Personal"
      End
      Begin VB.Menu registrados 
         Caption         =   "Registrados"
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "AYUDA"
      Begin VB.Menu sistema 
         Caption         =   "Sistema "
      End
   End
End
Attribute VB_Name = "frmpresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backup_Click()
frmbackup.Show
End Sub

Private Sub boletas_Click()
frmbol.Show
End Sub

Private Sub boletas1_Click()
frmboleta.Show
End Sub

Private Sub cliente_Click()
frmclientes.Show
End Sub

Private Sub clientes_Click()
dtacliente.Show
End Sub

Private Sub clientes1_Click()
frmcli.Show
End Sub

Private Sub clientes2_Click()
frmgrafico.Show
End Sub


Private Sub distrito_Click()
frmcdistrito.Show
End Sub

Private Sub distritos_Click()
frmdis.Show
End Sub

Private Sub factura_Click()
frmfactura.Show
End Sub

Private Sub facturas_Click()
frmfac.Show
End Sub

Private Sub facturasporfechas_Click()
frmgrafico.Show
End Sub

Private Sub Form_Load()


frmpresentacion.lbltitulo = frmreg.txtDNI(3)

If lbltitulo.Caption = "ALMACEN" Then
habilitacontroles

Else
deshabilitacontroles

End If
End Sub

Private Sub guiainterna_Click()
frmguia.Show
End Sub


Private Sub forma_Click()
dtafor.Show
End Sub

Private Sub ingresodedocumento_Click()
frmingreso.Show
End Sub

Private Sub lineasr_Click()
dtalineas.Show
End Sub

Private Sub Marca_Click()
frmcmarcas.Show
End Sub

Private Sub marcas_Click()
frmmar.Show
End Sub

Private Sub personas_Click()
frmper.Show
End Sub

Private Sub OrdendePedido1_Click()
frmordped.Show
End Sub

Private Sub marcasr_Click()
dtamarcas.Show
End Sub

Private Sub notaguiainterna_Click()
frmguia.Show
End Sub

Private Sub notapedidointerno_Click()
frmpedido.Show
End Sub

Private Sub OrdenCompra_Click()
frm_comp.Show
End Sub

Private Sub pagocuenta_Click()
frmpago.Show
End Sub

Private Sub pagosacuenta_Click()
frmpagoc.Show
End Sub

Private Sub personal_Click()
frmper.Show
End Sub


Private Sub producto_Click()
dtapersonas.Show
End Sub

Private Sub personales_Click()
dtapersonas.Show
End Sub

Private Sub productos_Click()
frmproducto.Show
End Sub

Private Sub Productos1_Click()
frmpro.Show
End Sub


Private Sub proveedores_Click()
frmproveedores.Show
End Sub

Private Sub proveedores1_Click()
frmprov.Show
End Sub

Private Sub registrados_Click()
frmregusu.Show
End Sub



Private Sub Timer1_Timer()
lbltitulo.ForeColor = QBColor(Int(Rnd * 15))
End Sub

Public Sub habilitacontroles()
proveedores1.Enabled = True
Productos1.Enabled = True
personal.Enabled = True
cargos.Enabled = True
marcas.Enabled = True
ingresodedocumento.Enabled = True
productos.Enabled = True
proveedores.Enabled = True
OrdenCompra.Enabled = True
notaguiainterna.Enabled = True

pagocuenta.Enabled = False
clientes1.Enabled = False
distritos.Enabled = False
factura.Enabled = False
boletas1.Enabled = False

cliente.Enabled = False
notapedidointerno.Enabled = False
End Sub
Public Sub deshabilitacontroles()
proveedores1.Enabled = False
Productos1.Enabled = False
personal.Enabled = False
cargos.Enabled = False
marcas.Enabled = False
ingresodedocumento.Enabled = False
productos.Enabled = False
proveedores.Enabled = False
OrdenCompra.Enabled = False
notaguiainterna.Enabled = False

pagocuenta.Enabled = True
clientes1.Enabled = True
distritos.Enabled = True
factura.Enabled = True
boletas1.Enabled = True


OrdendePedido.Enabled = True
cliente.Enabled = True
notapedidointerno.Enabled = True

End Sub

Private Sub Timer2_Timer()
Randomize
a = Int((Rnd * 10) + 1)
b = Int((Rnd * 10) + 1)


Image3.Left = a * 800
Image3.Top = b * 800 + 200

Image4.Left = a * 400
Image4.Top = b * 400 + 200





End Sub


Private Sub Timer3_Timer()
Image01.Visible = Not Image01.Visible
Image02.Visible = Not Image02.Visible
End Sub

Private Sub VentaporVendedorporFechas_Click()

End Sub

Private Sub vehiculos_Click()
frmveh.Show
End Sub
