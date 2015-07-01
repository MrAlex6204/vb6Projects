VERSION 5.00
Begin VB.Form frmEliminar 
   BackColor       =   &H00000000&
   Caption         =   "Baja de usuarios"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6945
   WindowState     =   2  'Maximized
   Begin VB.ListBox ListUsuarios 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2340
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   11655
   End
   Begin VB.Frame frmeMostrar 
      BackColor       =   &H00000000&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   6855
      Begin VB.CommandButton Command3 
         Caption         =   "C&errar"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblnom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblape 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lbldirec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lbltel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono.:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre(s):"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   3000
      Picture         =   "frmEliminar.frx":0000
      Top             =   0
      Width           =   6435
   End
End
Attribute VB_Name = "frmEliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
ListUsuarios.Appearance = 0

Dim ArchTemp As Integer
Dim Nom, Ape, Direc, Tel As String

ArchUsuarios = FreeFile
ArchTemp = 15

If Verificar_Existe(App.path + "\ArchMaster.Dat") = True Then


Open App.path + "\ArchMaster.Dat" For Input As #ArchUsuarios
Open App.path + "\Temp.bkup" For Output As #ArchTemp


'Crea el El archivo Teporal Con los Datos nuevos
Do While Not EOF(ArchUsuarios)
    
        Line Input #ArchUsuarios, Nom
        Line Input #ArchUsuarios, Ape
        Line Input #ArchUsuarios, Direc
        Line Input #ArchUsuarios, Tel
        
'Si es Diferente pasa el registro
'al archivo Temp.txt de lo contrario no
'y continua leyendo el registro
'Nota para poder eliminar el registro
'es necesario que todos los datos coincidan
If Nom <> lblnom And Ape <> lblape And Direc <> lbldirec And Tel <> lbltel Then
Print #ArchTemp, UCase(Nom)
Print #ArchTemp, UCase(Ape)
Print #ArchTemp, UCase(Direc)
Print #ArchTemp, UCase(Tel)

End If
        
Loop
'se cierran todos los Archivos Abieros
Close #ArchTemp
Close #ArchUsuarios



Open App.path + "\ArchMaster.Dat" For Output As #ArchUsuarios
Open App.path + "\Temp.bkup" For Input As #ArchTemp

'Actualiza el Archivo Usuarios
Do While Not EOF(ArchTemp)

        'Lee los datos de Temp.bkup y escribe en el
        'Archivo ArchMaster.Dat
        Line Input #ArchTemp, Nom
        Line Input #ArchTemp, Ape
        Line Input #ArchTemp, Direc
        Line Input #ArchTemp, Tel
                
        'Escribe los Datos leidos en el archivo uusario.txt
        
        Print #ArchUsuarios, UCase(Nom)
        Print #ArchUsuarios, UCase(Ape)
        Print #ArchUsuarios, UCase(Direc)
        Print #ArchUsuarios, UCase(Tel)
        
        

Loop
Close #ArchTemp
Close #ArchUsuarios
Command1.Enabled = True

ListUsuarios.Clear
Call CargaDatos
MsgBox "Usuario Eliminado", vbCritical, "VeraSoft Development"
Else
MsgBox "No se Encontro el Archivo: ArchMaster.Dat", vbCritical, "VeraSoft Development"
Me.Hide
End If




End Sub

Sub CargaDatos()
ListUsuarios.Appearance = 0
ArchUsuarios = FreeFile


If Verificar_Existe(App.path + "\ArchMaster.Dat") = True Then

Open App.path + "\ArchMaster.Dat" For Input As #ArchUsuarios

Dim Contenido As String
Do While Not EOF(ArchUsuarios)
        
        'Lee la linea
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "Nombre(s):" + Contenido
        
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Apellidos:" + Contenido
        
        Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Direccion:" + Contenido
        
         Line Input #ArchUsuarios, Contenido
        ListUsuarios.AddItem "   Telefono :" + Contenido
        
        ListUsuarios.AddItem ""
Loop
Close #ArchUsuarios

Else
MsgBox "No se Encontro el Archivo: ArchMaster.Dat", vbCritical, "VeraSoft Development"
Me.Hide
End If




End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call CargaDatos
End Sub

Private Sub ListUsuarios_Click()
frmeMostrar.Visible = True
Dim cadena As String
Dim i As Integer
If ListUsuarios.ListIndex <> -1 Then
   'para poner el dato de un ittem seleccionado
  cadena = Mid(ListUsuarios.List(ListUsuarios.ListIndex), 1, 10)
  If cadena = "Nombre(s):" Then
  lblnom.Visible = True
  lblape.Visible = True
  lbldirec.Visible = True
  lbltel.Visible = True
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex))
  lblnom = Mid(ListUsuarios.List(ListUsuarios.ListIndex), 11, (i - 10 + 1))
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 1))
  lblape = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 1), 14, (i - 10 + 1))
  
   i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 2))
  lbldirec = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 2), 14, (i - 10 + 1))
  
  i = Len(ListUsuarios.List(ListUsuarios.ListIndex + 3))
  lbltel = Mid(ListUsuarios.List(ListUsuarios.ListIndex + 3), 14, (i - 10 + 1))
  
  End If

End If
End Sub

