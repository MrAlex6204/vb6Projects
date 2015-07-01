Attribute VB_Name = "Varios"
Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByRef lParam As Any) As Long
                   
Public Type POINTAPI
    X As Long
    y As Long
End Type
Private Sign(255) As Integer
Private Const EM_CHARFROMPOS As Long = &HD7&
'---Comprovar si existe

'url
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
'Archivo
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Function GetWord(Rich As RichTextBox, ByVal X&, ByVal y&) As String
    Dim pos As Long, P1 As Long, P2 As Long
    Dim Char As Long
    Dim MousePointer As POINTAPI
    
    ' Position des Textzeichens unter dem Mauszeiger auslesen.
    MousePointer.X = X \ Screen.TwipsPerPixelX
    MousePointer.y = y \ Screen.TwipsPerPixelY
    pos = SendMessage(Rich.hWnd, EM_CHARFROMPOS, 0&, MousePointer)
    If pos <= 0 Then Exit Function
    
    ' Wortanfang finden.
    For P1 = pos To 1 Step -1
        Char = Asc(Mid$(Rich.Text, P1, 1))
        If Sign(Char) = 2 Then
            Exit For
        End If
    Next P1
    P1 = P1 + 1
    
    ' Wortende finden.
    For P2 = pos To Len(Rich.Text)
        Char = Asc(Mid$(Rich.Text, P2, 1))
        If Sign(Char) = 2 Then
            Exit For
        End If
    Next P2
    P2 = P2 - 1
    
    If P1 < P2 Then GetWord = Mid$(Rich.Text, P1, P2 - P1 + 1)
End Function
Public Function InitSigns()
    Dim i As Long
    Dim k As Long
    Dim Test As String
    
    Test = ".,;:?!"
    For i = 1 To Len(Test)
        k = Asc(Mid$(Test, i, 1))
        Sign(k) = 1
    Next i
    
    Test = " " + vbCrLf + Chr$(160)
    For i = 1 To Len(Test)
        k = Asc(Mid$(Test, i, 1))
        Sign(k) = 2
    Next i
End Function
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function Existe(Ruta As String) As Boolean
On Error Resume Next
If CStr(CBool(PathFileExists(Ruta))) = True Or CStr(CBool(PathIsURL(Ruta))) = True _
Or LCase(Left(Ruta, 4)) = "www." Or LCase(Left(Ruta, 4)) = "ftp." Then
Existe = True
End If
End Function

Public Function Remplazar(Cadena As String, flag As Boolean) As String
Dim CadenaA  As String, CadenaB As String, i As Integer
CadenaA = "\/:*?" & Chr(34) & "<>|"
CadenaB = "ΆΨΓ£λυγ₯"
If flag Then
For i = 1 To 9
Cadena = Replace(Cadena, Mid(CadenaA, i, 1), Mid(CadenaB, i, 1))
Next
Else
For i = 1 To 9
Cadena = Replace(Cadena, Mid(CadenaB, i, 1), Mid(CadenaA, i, 1))
Next
End If
Remplazar = Cadena
End Function
