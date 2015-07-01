Attribute VB_Name = "Module1"
Option Explicit

Private Const BIF_BROWSEFORCOMPUTER = 1000
Private Const BIF_BROWSEFORPRINTER = 2000
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_RETURNFSANCESTORS = 8
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4

Private Const MAX_SIZE = 255

Private Type BROWSEINFO
         hwndOwner As Long
         pidlRoot As Long
         pszDisplayName As String
         lpszTitle As String
         ulFlags As Long
         lpfn As Long
         lParam As Long
         iImage As Long
End Type

Private Declare Function BrowseFolderDlg Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function GetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal PointerToIDList As Long, ByVal pszPath As String) As Long

'Abre la ventana de seleccion de directorio:
Public Function DLG_BrowseFolder(hwnd As Long, Title As String) As String
         On Local Error Resume Next

         Dim mBrowseInfo As BROWSEINFO
         Dim mPointerToIDList As Long
         Dim mResult As Long
         Dim mPathBuffer As String
         Dim sReturn As String

         sReturn = vbNullString
         With mBrowseInfo
                      .hwndOwner = hwnd
                      .pidlRoot = 0
                      .lpszTitle = Title
                      .pszDisplayName = String(MAX_SIZE, Chr(0))
                      .ulFlags = BIF_RETURNONLYFSDIRS

         End With

         mPointerToIDList = BrowseFolderDlg(mBrowseInfo)

         If mPointerToIDList <> 0& Then
                      mPathBuffer = String(MAX_SIZE, Chr(0))
                      mResult = GetPathFromIDList(ByVal mPointerToIDList, ByVal mPathBuffer)
                      sReturn = Left$(mPathBuffer, InStr(mPathBuffer, Chr(0)) - 1)

         End If
 
         DLG_BrowseFolder = sReturn

End Function
