VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Tapiz ejemplo"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'función Api SystemParametersInfo para establecer el papel Tapiz de windows
Private Declare Function SystemParametersInfo _
    Lib "user32" _
    Alias "SystemParametersInfoA" ( _
        ByVal uAction As Long, _
        ByVal uParam As Long, _
        ByVal lpvParam As String, _
        ByVal fuWinIni As Long) As Long

'Variable para escribir los cambios en el registro con Windows Scriptin Host
Dim Wsh As Object
'Variable para obtener el directorio del sistema
Dim Fso As Object

Private Sub cambiarTapiz(Ruta As String, opcion As Byte)
'Usamos Windows Scriptin Host y Windows Scripting runtime
Set Wsh = CreateObject("wscript.shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

'Dependiendo de la opcion grabamos en el registro la configuración del tapiz

Select Case opcion
    'Papel tapiz centrado
    Case 0
        Wsh.regwrite "HKCU\Control Panel\Desktop\WallpaperStyle", "0"
        Wsh.regwrite "HKCU\Control Panel\Desktop\TileWallpaper", "0"
    'Papel tapiz - Mosaico
    Case 1
        Wsh.regwrite "HKCU\Control Panel\Desktop\WallpaperStyle", "0"
        Wsh.regwrite "HKCU\Control Panel\Desktop\TileWallpaper", "1"
    'Papel tapiz - expandida o estirado
    Case 2
        Wsh.regwrite "HKCU\Control Panel\Desktop\WallpaperStyle", "2"
        Wsh.regwrite "HKCU\Control Panel\Desktop\TileWallpaper", "0"
End Select

'Grabamos en en el directorio del sistema la imagen con savePicture
SavePicture LoadPicture(Ruta), Fso.GetSpecialFolder(1) & "\Imagen.bmp"

'Establecemos la imagen con SystemParametersInfo
SystemParametersInfo 20, 0&, Fso.GetSpecialFolder(1) & "\Imagen.bmp", &H1 Or &H2

'Eliminamos las variables de objeto
Set Wsh = Nothing
Set Fso = Nothing
End Sub





Private Sub Command1_Click()
With CommonDialog1
        .Filter = "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif"
        .ShowOpen

        If .FileName = "" Then Exit Sub
        
    End With
    'Le pasamos la ruta y el modo en que se establece el papel tapiz.
        'el 0 es centrado, el 1 es mosaico y el 2 es extendido
        cambiarTapiz CommonDialog1.FileName, 0
End Sub
