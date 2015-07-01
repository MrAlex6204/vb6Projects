Attribute VB_Name = "OPEN_FILENAME"
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Dim OFName As OPENFILENAME
'---------Carpetas guardar
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
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Function Guardar() As String
    
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        
        .hWndOwner = Form1.hWnd
        .lpszTitle = lstrcat("C:\", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    Guardar = sPath
End Function



Public Function ShowOpen(Extenciones As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = Form1.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Extenciones
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    'OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Abrir"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function

Public Function ShowSave() As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = Form1.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Todos los archivos (*.*)" + Chr$(0) + "*.*"
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = "C:\"
    OFName.lpstrTitle = "Guardar"
    OFName.flags = 0

    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function



