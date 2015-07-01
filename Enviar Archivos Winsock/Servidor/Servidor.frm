VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Servidor"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Poner en Escucha"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Dierectorio donde guarda le archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Puerto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para el cuadro de diálogo Seleccionar carpeta de windows
'*********************************************************
Private Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'**********************************************************

Dim Flag As Boolean
Dim sizeFileRecibido As Long
Dim sizeFile As Long

Private Sub Command1_Click()
On Error Resume Next
'Le asignanmos el número de puerto
Winsock1.LocalPort = Text1
'Ponemos a la escucha
Winsock1.Listen

Command1.Enabled = False
End Sub

Private Sub Command2_Click()
'Mostramos en el Text2 la ruta donde se guardará el archivo
Text2 = Ruta(Me)
End Sub

Private Sub Form_Load()
Text1 = "3000"
Text2.Enabled = False
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next

Command1.Enabled = False
'Cerramos el Winsock
Winsock1.Close
'Aceptamos la conexión del Winsock2
Winsock2.Accept requestID

MsgBox "Conexion recibida"
End Sub


Private Sub Winsock2_Close()
On Error Resume Next

Command1.Enabled = True

Winsock2.Close

MsgBox "Conexion cerrada"
End Sub

Private Sub Winsock2_Connect()
MsgBox "Conexion aceptada"
End Sub


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

'Array de Bytes para escribir el archivo en disco
Dim Archivo() As Byte

If Flag = False Then
Winsock2.GetData Datos, vbString
If Mid(Datos, 1, 9) = "|Archivo|" Then
' Flag
Flag = True
'Ponemos en 0
sizeFileRecibido = 0
' Separamos los datos
Datos = Split(Datos, "|")

sizeFile = Datos(2)
'Ponemos el ProgressBar en 0
ProgressBar1.Value = 0
'Establecemos el Max del ProgressBar pasandole comomáximo el tamaño del archivo
ProgressBar1.Max = sizeFile
' Le enviamos como mensaje al cliente que comienze el envio del archivo
Winsock2.SendData "|Ok|"

'Creamos un archivo en modo binario pasandole la ruta del text2
Open Text2 & "\" & Datos(3) For Binary Access Write As #1
End If
End If

If Flag = True Then

' Aumentamos sizeFileRecibido con los datos que van llegando
sizeFileRecibido = sizeFileRecibido + bytesTotal

'Recibimos los datos y lo almacenamos en el arry de bytes
Winsock2.GetData Archivo

'Colocamos en el valor de lo recibido en el Value del progressbar
ProgressBar1.Value = sizeFileRecibido

'Escribimos en disco el array de bytes, es decir lo que va llegando
Put #1, , Archivo

' Si lo recibido es mayor o igual al tamaño entonces se terminó y cerramos
'el archivo abierto
If sizeFileRecibido >= sizeFile Then
'Cerramos el archivo
Close #1
'Reestablecemos el flag y la variable sizeFileRecibido por si se intenta enviar otro archivo
Flag = False
sizeFileRecibido = 0
'Actualizar dato del ProgressBar
ProgressBar1.Value = ProgressBar1.Max
'Mostrar mensaje de finalización
MsgBox "Archivo se ha recibido por completo"
End If
End If

End Sub


Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

Command1.Enabled = True
'Cerramos el Winsock
Winsock2.Close
'Mostramos el aviso de que se cerró la conexión
MsgBox "La Conexion se ha cerrado", vbInformation
End Sub


'Función para abrir el cuadro de dialogo de windows y retornar el path que
'se visualiza en el text2
'***********************************************************************
Private Function Ruta(f As Form) As String

Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

With udtBI

.hWndOwner = f.hWnd
.lpszTitle = lstrcat("C:\", "")

.ulFlags = BIF_RETURNONLYFSDIRS
End With

'Mostramos el cuadro de diálogo "Buscar carpeta de windows"
lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
sPath = String$(MAX_PATH, 0)
'Get the path from the IDList
SHGetPathFromIDList lpIDList, sPath

CoTaskMemFree lpIDList
iNull = InStr(sPath, vbNullChar)
If iNull Then
sPath = Left$(sPath, iNull - 1)
End If
End If
'Retornamos el Path a la función que luego se muestra en el text2
Ruta = sPath
End Function
