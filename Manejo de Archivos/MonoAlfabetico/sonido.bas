Attribute VB_Name = "sonido"
Option Explicit

Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Dim ret As Long, path As String 'Api para reproducir sonidos

Sub CargarSonido(Pat As String) ' Pat = ruta del archivo temporal
Dim myArray() As Byte
Dim myFile As Long
myArray = LoadResData(102, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile
Open Pat For Binary Access Write As #myFile
Put #myFile, , myArray ' Escribe el archivo temporal
Close #myFile
End Sub

Private Sub Load()
Call CargarSonido("c:\sonido. mp3") ' Llama a la funcion q crea el archivo temporal
Call mciExecute("open " & "c:\sonido. mp3") ' Reproduce el archivo temporal
End Sub

Private Sub Unload(Cancel As Integer)
mciExecute "Close All" 'Detiene la reproduccion
Kill "c:\sonido.mp3" ' elimina el archivo temporal

End Sub


