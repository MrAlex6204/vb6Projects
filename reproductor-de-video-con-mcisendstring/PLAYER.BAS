Attribute VB_Name = "PLAYER1"
'Codigo Solo 16 Bits(vb2, vb3, vb4 de 16)
'*** Constantes Globales ***
'Global Const OFN_FILEMUSTEXIST = &H1000&
'Global Const OFN_READONLY = &H4&


'*** Variables Globales ***
'Global DialogCaption As String
'Global FileName As String

'*** Declaraciones de función Globales ***
'Declare Function mciGetErrorString Lib "mmsystem" (ByVal MCI_ERROR As Long, ByVal ErororString As String, ByVal ReturnLenght As Integer) As Integer
'Declare Function mciSendString Lib "mmsystem" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen%, ByVal hCallBack%) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Const OFN_FILEMUSTEXIST = &H1000&
Public Const OFN_READONLY = &H4&
Public DialogCaption As String
Public FileName As String

Public Declare Function GetShortPathName Lib "kernel32" _
Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Temp As Variant
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

