VERSION 5.00
Begin VB.Form frmRepCajeros 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REPORTE DE CAJEROS"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   7920
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2640
      TabIndex        =   2
      Top             =   8160
      Width           =   1860
   End
   Begin VB.ListBox Listart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   9000
      Picture         =   "frmRepCajeros.frx":0000
      Top             =   3000
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "REPORTE DE CAJEROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   540
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   5475
   End
End
Attribute VB_Name = "frmRepCajeros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim num, nom As String
frmArtVendidos.Caption = "REPORTE DE ARTICULOS VENDIDOS DE: " + CajeroOnline
Listart.Clear
ind = 1
Open App.path + "\Cajeros.Dat" For Input As #33
            
            Do While Not EOF(33)
     
            Input #33, num, nom
            Listart.AddItem ind & "  NOMBRE:" & nom & "     NUM. CAJERO:" & num
            ind = ind + 1
            
Loop
Close #33
End Sub
