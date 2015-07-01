VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir diálogo"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Opción para visualizar el botón de ""Crear nueva Carpeta"""
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
 
' Estructura BrowseInfo requerida para el Api SHBrowseForFolder

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

' Constantes
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260 ' Para Buffer de caracteres del path

' Funcion Api CoTaskMemFree
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

' Funcion Api CoTaskMemFree lstrcat
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" ( _
    ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long

' Funcion Api SHBrowseForFolder
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

' Funcion Api SHGetPathFromIDList
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long


Private Sub Command1_Click()

Dim ret As Long
Dim sPath As String
Dim tBI As BrowseInfo

    With tBI
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat("C:\", "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or 64

    End With

    'Mostrar el cuadro de dialogo Buscar carpeta
    ret = SHBrowseForFolder(tBI)

    If ret Then
        
        sPath = String$(MAX_PATH, 0)
        Call SHGetPathFromIDList(ret, sPath)
        Call CoTaskMemFree(ret)
        
        Dim pos As Long
        pos = InStr(sPath, vbNullChar)

        If pos Then
            sPath = Left$(sPath, pos - 1)
        End If
    End If
    
    'Mostramos el Path elegido
    MsgBox " El directorio seleccionado es: " & sPath, vbInformation

End Sub

