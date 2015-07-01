Attribute VB_Name = "Declaraciones"
Public ABC(27), ENCRIP(26) As String
Public EncriptacionText, DesencriptacionText As String
Sub Main()




'El Es para verificar si existen los componentes necesarios para la Ejecucion
'del Prg

Dim myArray(), FuenteArray() As Byte
Dim myFile As Long



Dim fso, Mensaje As Object
ArchivoExiste = True


Dim path As String
path = Environ("SystemRoot") 'Me Devuelve El directorio Raiz c:\Windows y se le asigna a
' la variable Path

Set fso = CreateObject("Scripting.FileSystemObject")




If Not (fso.fileexists(path + "\system32\comctl32.ocx")) Then
myArray = LoadResData(103, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile
Open path + "\system32\comctl32.ocx" For Binary Access Write As #myFile
Put #myFile, , myArray ' Escribe el archivo temporal
Close #myFile
End If




'Descomprime la Fuente y despues la Agrega
FuenteArray = LoadResData(104, "CUSTOM") 'Carga el archivo de recursos
myFile = FreeFile
Open "C:\punk kid.ttf" For Binary Access Write As #myFile
Put #myFile, , FuenteArray ' Escribe el archivo Binario
Close #myFile



'Agrega la Fuente
Fuente.AgregarFuente ("C:\punk kid.ttf")



Load VeraSoft
VeraSoft.Show

End Sub
Public Sub AlfabetoLoad()
ABC(0) = "A"
ABC(1) = "B"
ABC(2) = "C"
ABC(3) = "D"
ABC(4) = "E"
ABC(5) = "F"
ABC(6) = "G"
ABC(7) = "H"
ABC(8) = "I"
ABC(9) = "J"
ABC(10) = "K"
ABC(11) = "L"
ABC(12) = "M"
ABC(13) = "N"
ABC(14) = "O"
ABC(15) = "P"
ABC(16) = "Q"
ABC(17) = "R"
ABC(18) = "S"
ABC(19) = "T"
ABC(20) = "U"
ABC(21) = "V"
ABC(22) = "W"
ABC(23) = "X"
ABC(24) = "Y"
ABC(25) = "Z"
End Sub
Sub Grabar(Cadena As String)


Open App.path + "\Encript.Dat" For Append As #10
Print #10, UCase(Cadena)

Close #10
End Sub

Sub Mensaje(Mensaje As String)

VeraSoft.FrameMensaje.Visible = True
VeraSoft.lblMensaje = Mensaje

End Sub
