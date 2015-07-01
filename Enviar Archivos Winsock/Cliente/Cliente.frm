VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Enviar Archivo"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   9135
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5280
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3840
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Puerto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   720
         TabIndex        =   9
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "_____________________________________________________________________"
         ForeColor       =   &H8000000C&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   6210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "ip de Servidor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   600
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
'conectamos al servidor. El Text1 es la dirección IP y el Text2 es el puerto
Winsock1.Connect Text1, Text2

Command1.Enabled = False

End Sub

Private Sub Command2_Click()
'Comprobamos que hay un archivo a enviar
If Trim(Text3) = "" Then
MsgBox "Debe elegir un archivo"
Exit Sub
End If

If Dir(Text3) <> "" Then
' Separar el nombre del archivo para solo tomar su nombre (sin la ruta)
Datos = Split(Text3, "\")

'Datos(UBound(Datos)) es el nombre del archivo, que sería Datos(3)

'Datos(2) es el tamaño en bytes del archivo. Esta información se la enviamos
'antes de enviar el fichero

'enviamos los datos
Winsock1.SendData "|Archivo|" & FileLen(Text3) & "|" & Datos(UBound(Datos))
Else
MsgBox "El archivo no existe"
End If
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
Text3 = CommonDialog1.FileName
End If
Command2.Enabled = True
End Sub

Private Sub Form_Load()
'Ip del formulario servidor
Text1 = "127.0.0.1"
'Puerto
Text2 = "3000"
Text3.Enabled = False
End Sub

Private Sub Winsock1_Close()
On Error Resume Next

Command1.Enabled = True
Command2.Enabled = False
'Cerramos el winsock
Winsock1.Close

MsgBox "La Conexion se ha cerradado"
End Sub

Private Sub Winsock1_Connect()
Command1.Enabled = False
Command2.Enabled = True

MsgBox "Conectado correctamente al servidor"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Datos, vbString


If Datos = "|Ok|" Then
Enviar_Archivo
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

Command1.Enabled = True
Command2.Enabled = False
'Cerramos el winsock
Winsock1.Close
MsgBox "Error en la conexion"
End Sub

Private Sub Enviar_Archivo()
Dim Size As Long
Dim Archivo() As Byte

Open Text3 For Binary Access Read As #1
'Obtenemos el tamaño exacto en bytes del archivo para poder redimensionar el array de bytes
Size = LOF(1)
ReDim Archivo(Size - 1)
'Leemos y almacenamos todo el fichero en el array
Get #1, , Archivo
'Cerramos
Close

'Enviamos el archivo
Winsock1.SendData Archivo
End Sub

