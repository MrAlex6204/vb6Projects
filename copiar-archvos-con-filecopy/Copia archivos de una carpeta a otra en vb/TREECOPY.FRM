VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TreeCopy 1.0"
   ClientHeight    =   3825
   ClientLeft      =   3105
   ClientTop       =   2895
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TREECOPY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   5910
   Begin VB.DirListBox Dir2 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.DriveListBox Drive2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Barra de estado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Directorio Destino:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Disco:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Directorio Fuente:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Disco:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu MnuFila 
      Caption         =   "&Fila"
      Begin VB.Menu MnuSale 
         Caption         =   "&Salir"
         Index           =   1
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "&Acciones"
      Begin VB.Menu MnuMover 
         Caption         =   "&Copiar a otro directorio"
         Index           =   2
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu MnuAcerca 
         Caption         =   "&Acerca de..."
         Index           =   3
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Sub Drive1_Change()
'Cambiar la unidad de disco                (08/Oct/06)
    On Error GoTo ErrorDeDisco
    Dir1.Path = Drive1.Drive
    Exit Sub
ErrorDeDisco:
    Drive1.Drive = Dir1.Path
    Exit Sub

End Sub

Private Sub Drive2_Change()
 'Cambiar la unidad de disco                (08/Oct/06)
    On Error GoTo ErrorDeDisco1
    Dir2.Path = Drive2.Drive
    Exit Sub
ErrorDeDisco1:
    Drive2.Drive = Dir2.Path
    Exit Sub

End Sub

Private Sub Form_Resize()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29 'Centra el formulario completamente


End Sub

Private Sub MnuAcerca_Click(Index As Integer)
Call ShellAbout(Me.hwnd, "TreeCopy 1.0", "Copyright 2007, Dj_Dexter", Me.Icon)
End Sub

Private Sub MnuMover_Click(Index As Integer)
Dim Ruta, Ruta1, x, z
If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + Dir2.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
  Ruta = Dir1.Path & ""
 Else
  Ruta = Dir1.Path & "\"
End If
If Right(Dir2.Path, 1) = "\" Then
  Ruta1 = Dir2.Path & ""
 Else
  Ruta1 = Dir2.Path & "\"
End If

y = Ruta
z = Ruta1
result = xfilecopy("" & y & "", "" & z & "", "*.*", Label1)
If Err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If

End Sub

Private Sub MnuSale_Click(Index As Integer)
End
End Sub

Private Function xfilecopy(origen$, destino$, archivo$, informa As Label)
' Copia varios archivos de una carpeta a otra
' Origen$= directorio de origen , terminado en "\"
' Destino$= directorio de destino , terminado en "\"
' archivo$= especificacion de archivos a copiar, con simb. comodin
' informa= un label en el que se muestra el progreso
'
' result = xfilecopy("c:\pat\", "h:\doc\", "*.exe", label1)
' copia todos los archivos exe de c:\pat en h:\doc
' muestra lo que esta haciendo en label1


Dim n, result, cuenta, pcent
' cuenta los archivos a copiar
cuenta = 0
n = Dir$(origen$ & archivo$)
While (n <> "")
 cuenta = cuenta + 1
 n = Dir$
Wend

' Copia
result = 0
n = Dir$(origen$ & archivo$)
On Error GoTo malxfilecopy
While (n <> "") And (result > -1)
 pcent = (result + 1) & "/" & cuenta & " "
 pcent = pcent & Format$(100 * result / cuenta, "#0.0") & "%"
 informa.Caption = pcent & " Copiando " & origen$ & n & " a " & destino$
 DoEvents

 FileCopy origen$ & n, destino$ & n
 result = result + 1
 n = Dir$
continuaxfilecopy:
Wend
informa.Caption = ""
xfilecopy = result
Exit Function

malxfilecopy:
 result = -1
 Resume continuaxfilecopy
End Function

