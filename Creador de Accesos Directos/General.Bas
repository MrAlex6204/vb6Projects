Attribute VB_Name = "General"
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Type SHITEMID
     cb As Long
     abID As Byte
End Type

Type ITEMIDLIST
     mkid As SHITEMID
End Type

Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Public Const NOERROR = 0

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000


