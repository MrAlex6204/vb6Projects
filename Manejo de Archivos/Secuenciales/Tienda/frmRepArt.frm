VERSION 5.00
Begin VB.Form frmRepArt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REPORTE DE ARTICULOS"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4080
      TabIndex        =   2
      Top             =   8520
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
      Top             =   1440
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   9360
      Picture         =   "frmRepArt.frx":0000
      Top             =   2640
      Width           =   7020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "REPORTE DE ARTICULOS"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5955
   End
End
Attribute VB_Name = "frmRepArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim CODE, PRECIO, DESCRIP As String
frmArtVendidos.Caption = "REPORTE DE ARTICULOS VENDIDOS DE: " + CajeroOnline
Listart.Clear
ind = 1
Open App.path + "\Articulos.Dat" For Input As #33
            
            Do While Not EOF(33)
     
            Input #33, CODE, PRECIO, DESCRIP
            Listart.AddItem ind & "<--CODE: " & CODE & " PRECIO:" & PRECIO & " DESCRIP: " & DESCRIP
            ind = ind + 1
            
Loop
Close #33
End Sub
