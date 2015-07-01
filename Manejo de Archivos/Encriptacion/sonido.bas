Attribute VB_Name = "sonido"
Option Explicit

Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Dim ret As Long, path As String 'Api para reproducir sonidos
Sub CargarSonido(Pat As String) ' Pat = ruta del archivo temporal
Dim myArray() As Byte
Dim myFile As Long
myArray = LoadResData(101, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile + 1
Open Pat For Binary Access Write As #myFile
Put #myFile, , myArray ' Escribe el archivo temporal
Close #myFile
End Sub



