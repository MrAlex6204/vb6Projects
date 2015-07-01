Attribute VB_Name = "Fuente"
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
'Todas las fuentes son de tipo .ttf

Sub AgregarFuente(Fuente As String)
Dim X As Long
X = AddFontResource(Fuente)
End Sub

