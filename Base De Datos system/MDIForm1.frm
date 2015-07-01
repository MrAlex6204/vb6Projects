VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{DB3F8F1D-3ADE-4D2C-BA1A-BACA667F0EE4}#1.0#0"; "SysTrayocx.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "Universidada Autonoma de Tamaulipas"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":030A
   StartUpPosition =   2  'CenterScreen
   Begin sysTray.Tray Tray1 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Base de Datos system"
      IconPicture     =   "MDIForm1.frx":15178
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   7805
            Text            =   "Universidad Autonoma de Tamaulipas"
            TextSave        =   "Universidad Autonoma de Tamaulipas"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "11/01/2004"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1588
            Picture         =   "MDIForm1.frx":15492
            TextSave        =   "22:33"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu CerrarSis 
      Caption         =   "&Salir"
      Begin VB.Menu Cerrar 
         Caption         =   "CerrarSistema"
      End
   End
   Begin VB.Menu reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu ingeneria 
         Caption         =   "&Ingeneria"
      End
      Begin VB.Menu medicina 
         Caption         =   "&Medicina"
      End
   End
   Begin VB.Menu consultas 
      Caption         =   "&Consultas"
      Begin VB.Menu ConSistemas 
         Caption         =   "&Sistemas"
      End
      Begin VB.Menu ConMedicina 
         Caption         =   "&Medicina"
      End
   End
   Begin VB.Menu min 
      Caption         =   "Ocultar"
   End
   Begin VB.Menu VerTodod 
      Caption         =   "Ver Todos"
      Begin VB.Menu VerSistemas 
         Caption         =   "Ssietmas"
      End
      Begin VB.Menu VerMedicnina 
         Caption         =   "Medicinia"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long


      Private Declare Sub ReleaseCapture Lib "User32" ()
     
      

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2


Private Sub AltaMed_Click()
frmMedicina.Show
frmMedicina.Text1.Enabled = True
frmMedicina.Text2.Enabled = True
frmMedicina.Text3.Enabled = True
frmMedicina.Text4.Enabled = True

frmMedicina.Command1.Visible = True
frmMedicina.Command2.Visible = True
frmMedicina.Data1.Recordset.AddNew

End Sub

Private Sub Baja_Click()
frmSistemas.Show
End Sub

Private Sub BajaMed_Click()
frmMedicina.Show
End Sub

Private Sub BuscarMed_Click()
frmMedicina.Show
End Sub

Private Sub Cerrar_Click()

End
End Sub

Private Sub ConMedicina_Click()
frmMedicina.Show

End Sub

Private Sub ConSistemas_Click()
frmSistemas.Show
End Sub

Private Sub Editar_Click()
frmSistemas.Show
End Sub

Private Sub EditarMed_Click()
frmMedicina.Show
End Sub

Private Sub Guardar_Click()
frmSistemas.Show
frmSistemas.Data1.UpdateRecord
frmSistemas.Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
End Sub

Private Sub GuardarMed_Click()
frmMedicina.Show
frmMedicina.Data1.UpdateRecord
frmMedicina.Data1.Refresh
MsgBox "El Registro ha sido Guardado en la Base de Datos", vbExclamation, "Aviso Importante"
End Sub

Private Sub ingeneria_Click()
On Error GoTo ErrorOpen
'El sig. Codigo es para que DataEnvironment1.Alumnos _
se Conecte con la base de datos el path es segun donde se encuentre al aplicacion _
la funcion App.Path devuelve el path donde se encuentra nuestro prog. _
la Base de Datos se Encuentaen una carpeta llamada MiBaseDeDatos _
y la Base de Datos se llama MiBaseDeDatos.mdb

DataEnvironment2.Sistemas.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Sistemas.mdb" + "; Persist Security Info=False"


'esta parte el cdigo Muestra Nuestro DataReport
DataReport2.Show
Exit Sub

'Esta Parte del Codigo es Para x Si Hay un Error Al Abrir el DataReport
ErrorOpen:

Unload DataReport2
DataReport2.Show

End Sub

Private Sub MDIForm_Load()
Transparent.Aplicar_Transparencia Me.hWnd, 245
Globo.Mensaje ("VeraSotf Develoment")

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngReturnValue As Long
 
        If Button = 1 Then
        Transparent.Aplicar_Transparencia Me.hWnd, 150
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&)
        Mover.MoverForm
        Else
        Transparent.Aplicar_Transparencia Me.hWnd, 240
        End If
        
Globo.Remover
End Sub

Private Sub medicina_Click()
On Error GoTo ErrorOpen
'El sig. Codigo es para que DataEnvironment1.Alumnos _
se Conecte con la base de datos el path es segun donde se encuentre al aplicacion _
la funcion App.Path devuelve el path donde se encuentra nuestro prog. _
la Base de Datos se Encuentaen una carpeta llamada MiBaseDeDatos _
y la Base de Datos se llama MiBaseDeDatos.mdb

DataEnvironment1.MEDICINA.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Medicina.mdb" + "; Persist Security Info=False"


'esta parte el cdigo Muestra Nuestro DataReport
DataReport1.Show
Exit Sub

'Esta Parte del Codigo es Para x Si Hay un Error Al Abrir el DataReport
ErrorOpen:

Unload DataReport1
DataReport1.Show


End Sub

Private Sub task_Click()

End Sub

Private Sub min_Click()
Tray1.PonerSystray
Me.Hide
End Sub

Private Sub Tray1_DblClick(Button As Integer)
If Button = vbLeftButton Then
Me.Show
Tray1.RemoverSystray
End If
End Sub

Private Sub VerSistemas_Click()
Form1.Show
End Sub
