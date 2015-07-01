Attribute VB_Name = "Tapiz"
Option Explicit
'función Api SystemParametersInfo para establecer el papel Tapiz de windows
'Api SystemParametersInfo
Public Declare Function SystemParametersInfo _
    Lib "user32" _
    Alias "SystemParametersInfoA" ( _
        ByVal uAction As Long, _
        ByVal uParam As Long, _
        ByVal lpvParam As Any, _
        ByVal fuWinIni As Long) As Long

'Constantes
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = &H1

'______________________OTRA FORMA DE CAMBIAR EL PAPEL TAPIZ_________________________________________________________

'Variable para escribir los cambios en el registro con Windows Scriptin Host
Dim Wsh As Object
'Variable para obtener el directorio del sistema
Dim Fso As Object

 Sub cambiarTapiz(Ruta As String, opcion As Byte)
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
'SystemParametersInfo 20, 0&, Fso.GetSpecialFolder(1) + "\Imagen.bmp", &H1 Or &H2

'Eliminamos las variables de objeto
Set Wsh = Nothing
Set Fso = Nothing
End Sub

