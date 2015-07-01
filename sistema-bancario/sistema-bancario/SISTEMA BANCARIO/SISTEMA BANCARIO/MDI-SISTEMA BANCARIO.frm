VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm MDISIS 
   BackColor       =   &H00404040&
   Caption         =   "SISTEMA BANCARIO"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI-SISTEMA BANCARIO.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2595
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Picture         =   "MDI-SISTEMA BANCARIO.frx":14584
            Text            =   "SISTEMA BANCARIO"
            TextSave        =   "SISTEMA BANCARIO"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2647
            MinWidth        =   2647
            Picture         =   "MDI-SISTEMA BANCARIO.frx":149D6
            TextSave        =   "24/05/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "MDI-SISTEMA BANCARIO.frx":14E28
            TextSave        =   "07:47 a.m."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Picture         =   "MDI-SISTEMA BANCARIO.frx":1527A
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Picture         =   "MDI-SISTEMA BANCARIO.frx":156CC
            TextSave        =   "MAYÚS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNUMANT 
      Caption         =   "&MANTENIMIENTO"
      Begin VB.Menu MNUCLI 
         Caption         =   "&CLIENTES"
      End
      Begin VB.Menu MNUCTA 
         Caption         =   "&CUENTAS"
      End
      Begin VB.Menu MNUEMP 
         Caption         =   "&EMPLEADOS "
      End
   End
   Begin VB.Menu MNUMOV 
      Caption         =   "&REGISTRAR"
      Begin VB.Menu MNUNOMV 
         Caption         =   "NUEVO MOVIMIENTO"
      End
   End
   Begin VB.Menu MNUCON 
      Caption         =   "&CONSULTAS"
      Begin VB.Menu MNUCCTA 
         Caption         =   "POR CLIENTES"
      End
      Begin VB.Menu MNUCE 
         Caption         =   "POR EMPLEADOS"
      End
      Begin VB.Menu MNUCC 
         Caption         =   "POR CUENTAS"
      End
      Begin VB.Menu MNUPM 
         Caption         =   "POR MOVIMIENTOS"
      End
   End
   Begin VB.Menu MNUDEST 
      Caption         =   "&DATOS ESTADISTICOS"
      Begin VB.Menu MNUMPM 
         Caption         =   "MOV. POR MESES"
      End
      Begin VB.Menu MNUOPB 
         Caption         =   "RESUMEN x TIPO DE CUENTA"
      End
      Begin VB.Menu MNURDIS 
         Caption         =   "RESUMEN x DISTRITOS"
      End
      Begin VB.Menu MNUCM 
         Caption         =   "CUENTAS x MESES"
      End
   End
   Begin VB.Menu MNUAYUDA 
      Caption         =   "&AYUDA"
      Begin VB.Menu MNUBD 
         Caption         =   "BUSQUEDA DE DATOS"
      End
   End
   Begin VB.Menu MNUUT 
      Caption         =   "&UTILIDADES"
      Begin VB.Menu MNUCALC 
         Caption         =   "CALCULADORA"
      End
      Begin VB.Menu MNUBN 
         Caption         =   "BLOCK DE NOTAS"
      End
   End
   Begin VB.Menu MNUSALIR 
      Caption         =   "&SALIR"
   End
End
Attribute VB_Name = "MDISIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Call BD
End Sub

Private Sub MNUBD_Click()
FRMBUS.Show
End Sub

Private Sub MNUBN_Click()
A = Shell("NOTEPAD.EXE", 1)
End Sub

Private Sub MNUCALC_Click()
X = Shell("CALC.EXE")
End Sub

Private Sub MNUCC_Click()
FRMCON3.Show
End Sub

Private Sub MNUCCTA_Click()
FRMCON1.Show
End Sub

Private Sub MNUCE_Click()
FRMCON2.Show
End Sub

Private Sub MNUCLI_Click()
FRMCLI.Show
End Sub

Private Sub MNUCM_Click()
FRMDATEST4.Show
End Sub

Private Sub MNUCTA_Click()
FRMCTA.Show
End Sub

Private Sub MNUEMP_Click()
FRMEMP.Show
End Sub

Private Sub MNUMPM_Click()
FRMDATEST1.Show
End Sub

Private Sub MNUNOMV_Click()
FRMOV.Show
End Sub

Private Sub MNUOPB_Click()
FRMDATEST2.Show
End Sub

Private Sub MNUPM_Click()
FRMCON4.Show
End Sub

Private Sub MNURDIS_Click()
FRMDATEST3.Show
End Sub

Private Sub MNUSALIR_Click()
If MsgBox("¿DESEA SALIR DEL SISTEMA?", vbExclamation + vbYesNo, "SISTEMA BANCARIO") = vbYes Then
   End
End If
End Sub

