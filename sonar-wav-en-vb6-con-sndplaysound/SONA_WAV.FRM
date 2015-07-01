VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproductor de Sonidos "
   ClientHeight    =   3045
   ClientLeft      =   180
   ClientTop       =   495
   ClientWidth     =   6870
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
   Icon            =   "SONA_WAV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   6870
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Acciones: Nombre Ruta/Sonar/Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Info/Salir:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1080
         TabIndex        =   10
         Top             =   1800
         Width           =   2895
         Begin VB.CommandButton Command6 
            Appearance      =   0  'Flat
            Caption         =   "Salir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1560
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4695
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   4200
         Picture         =   "SONA_WAV.frx":030A
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Pequeño reproductor de sonido, permite sonarlos, detenerlos, o hacelos sonar eternamente hasta pararlos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         TabIndex        =   9
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "(c)2007 Dj_Dexter Gato_cosmico16@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Nombre de Ruta a sonar el Wave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   240
         Picture         =   "SONA_WAV.frx":0DC4
         Top             =   2040
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Acciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Detener"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Sonar eter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Sonar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Abrir..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function sndPlaySound Lib "mmsystem.dll" (ByVal lpszSoundName As String, ByVal uFlags As Integer) As Long
'En vb de win32 5 o 6 usar:
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Type OPENFILENAME
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
Dim Temp As Variant
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' Constantes para los flags
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  ver asociación de archivo, si lo hay
Const SND_APPLICATION = &H80
'  nombre en la entrada de WIN.INI [sounds]
Const SND_ALIAS = &H10000
'  Nombre de entrada identificada a WIN.INI es [sounds]
Const SND_ALIAS_ID = &H110000
'  Sonar Asincronadamente ¡Ni usar el vb te dará error de desbordamiento!
Const SND_ASYNC = &H1
'  Sonar asincronadamente (defecto)
Const SND_SYNC = &H0
'  Sonar sincronadamente ¡No recomendado, pues el mci no te devuelve el control hasta
'  que termine de sonar el wave y desbordar la aplicación, mejor SND_ASYNC!
'  nombre es una fila
Const SND_FILENAME = &H20000
'  lpszSoundName a una fila de memoria ¡Tampoco usar el vb te dará error de desbordamiento!
Const SND_LOOP = &H8
'  repetir sonido eternamente al retorno de la función sndPlaySound
Const SND_MEMORY = &H4
'  silencio no defecto, si no hay sonido encontrado
Const SND_NODEFAULT = &H2
'  no detener fila que ya está sonado
Const SND_NOSTOP = &H10
'  no se inicia si el dispositivo está vacio
Const SND_NOWAIT = &H2000
'  purgar a eventos-no estáticos en la tarea asignada
Const SND_PURGE = &H40
'  nombre es un nombre de recurso .res
Const SND_RESOURCE = &H40004
'*** Constantes ***
Const OFN_FILEMUSTEXIST = &H1000&
Const OFN_READONLY = &H4&


'*** Variables ***
Dim DialogCaption As String
Dim Filename As String

Private Sub Command1_Click()
Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Form1.hWnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "Archivos wav" + Chr$(0) + "*.wav"
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = CurDir
ofn.flags = &H4
ofn.lpstrTitle = " Seleccionar archivo de audio"

Temp = ""
Temp = GetOpenFileName(ofn)



If Temp Then 'Si se ha elegido un fichero...
    Dim Fichero As String
    Fichero = Trim$(ofn.lpstrFile)
    Text1.Text = Fichero


Else 'Si se le ha dado a ESC...
End If

End Sub

Private Sub Command2_Click()
'Para tocar un WAV, lo llamas así:
Y = Text1.Text
X = sndPlaySound("" & Y & "", SND_ASYNC)

End Sub

Private Sub Command3_Click()
'Para detener lo que se esté tocando
z = Text1.Text
X = sndPlaySound(ByVal "", 15)

End Sub

Private Sub Command4_Click()
'Para tocar un WAV de forma repetitiva, lo llamas así:
Y = Text1.Text
X = sndPlaySound("" & Y & "", SND_ASYNC + SND_LOOP)

End Sub

Private Sub Command5_Click()
On Error Resume Next
 Filename = Text1.Text
    Memoria& = FileLen(Filename)
    MsgBox "Fila: " + Text1.Text + " Peso: " + m$ & Format$(Memoria& \ 1024, "###,###") & " KB", 64, "Info Wave"
If Err Then Beep
End Sub

Private Sub Command6_Click()
'Para detener lo que se esté tocando
z = Text1.Text
X = sndPlaySound(ByVal "", 15)
End

End Sub

Private Sub Form_Resize()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Para detener lo que se esté tocando
z = Text1.Text
X = sndPlaySound(ByVal "", 15)
End
End Sub

