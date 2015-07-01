Attribute VB_Name = "Declaraciones"
Public ENCRIP(28), DESENCRIP(28) As String
Sub Main()

'El Es para verificar si existen los componentes necesarios para la Ejecucion
'del Prg

Dim myArray(), FuenteArray() As Byte
Dim myFile As Long



Dim fso, Mensaje
ArchivoExiste = True

Dim path As String
path = Environ("SystemRoot")

'Descomprime la Libreria en Caso de Que No exista la Misma
Set fso = CreateObject("Scripting.FileSystemObject")

If Not (fso.FileExists(path + "\system32\comctl32.ocx")) Then
myArray = LoadResData(102, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile
Open path + "\system32\comctl32.ocx" For Binary Access Write As #myFile
Put #myFile, , myArray ' Escribe el archivo temporal
Close #myFile
End If





'Descomprime la Fuente y despues la Agrega
FuenteArray = LoadResData(103, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile
Open App.path + "\MadSience.ttf" For Binary Access Write As #myFile
Put #myFile, , FuenteArray ' Escribe el archivo Binario
Close #myFile

'Agrega la Fuente
Fuente.AgregarFuente ("MadSience.ttf")



Load Form1
Form1.Show
Form1.Refresh
End Sub
Public Sub ENCRIPDAT()
ENCRIP(0) = "A"
ENCRIP(1) = "B"
ENCRIP(2) = "C"
ENCRIP(3) = "D"
ENCRIP(4) = "E"
ENCRIP(5) = "F"
ENCRIP(6) = "G"
ENCRIP(7) = "H"
ENCRIP(8) = "I"
ENCRIP(9) = "J"
ENCRIP(10) = "K"
ENCRIP(11) = "L"
ENCRIP(12) = "M"
ENCRIP(13) = "N"
ENCRIP(14) = "Ñ"
ENCRIP(15) = "O"
ENCRIP(16) = "P"
ENCRIP(17) = "Q"
ENCRIP(18) = "R"
ENCRIP(19) = "S"
ENCRIP(20) = "T"
ENCRIP(21) = "U"
ENCRIP(22) = "V"
ENCRIP(23) = "W"
ENCRIP(24) = "X"
ENCRIP(25) = "Y"
ENCRIP(26) = "Z"
ENCRIP(27) = " "
End Sub
Public Sub DESENCRIPDAT()
DESENCRIP(0) = "D"
DESENCRIP(1) = "E"
DESENCRIP(2) = "F"
DESENCRIP(3) = "G"
DESENCRIP(4) = "H"
DESENCRIP(5) = "I"
DESENCRIP(6) = "J"
DESENCRIP(7) = "K"
DESENCRIP(8) = "L"
DESENCRIP(9) = "M"
DESENCRIP(10) = "N"
DESENCRIP(11) = "Ñ"
DESENCRIP(12) = "O"
DESENCRIP(13) = "P"
DESENCRIP(14) = "Q"
DESENCRIP(15) = "R"
DESENCRIP(16) = "S"
DESENCRIP(17) = "T"
DESENCRIP(18) = "U"
DESENCRIP(19) = "V"
DESENCRIP(20) = "W"
DESENCRIP(21) = "X"
DESENCRIP(22) = "Y"
DESENCRIP(23) = "Z"
DESENCRIP(24) = "A"
DESENCRIP(25) = "B"
DESENCRIP(26) = "C"
DESENCRIP(27) = "-"
End Sub

