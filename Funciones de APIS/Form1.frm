VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11610
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Examimar"
      Height          =   735
      Left            =   6600
      TabIndex        =   1
      Top             =   6000
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9763
      _Version        =   393217
      BackColor       =   255
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":2832
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Abre el cuadro de diálogo para seleccionar el archivo dll / exe etc..
Private Sub Command1_Click()
    With CommonDialog1
        
        .FileName = vbNullString
        .Filter = "Archivos dll|*.dll|Archivos exe|*.exe|Todos|*.*"
        .ShowOpen
        
        If .FileName = vbNullString Then Exit Sub
        
        Call Exportar_funciones(.FileName)
        
    End With
End Sub

' Exporta las funciones al archivo txt
Function Exportar_funciones(Path_Dll As String)
    
    Dim Path_txt As String
    Dim Nombre_Archivo As String
    Dim Pausa As Long
    Dim ret As Long
    Dim linkpath As String
    linkpath = App.Path
    ' Comadno del Link.exe para exportar las funciones
    Const Comando_Export As String = " /dump /exports "
    
    'Path del archivo Link.exe ( ubicado en el directorio _
     de instalación de visual basic )
    Dim Path_Link As String
    Path_Link = linkpath + "\LINK.EXE"
    
    ' Extrae el nombre del archivo Dll separándolo del path
    Nombre_Archivo = Right(Path_Dll, Len(Path_Dll) - InStrRev(Path_Dll, "\"))
    Me.Caption = Nombre_Archivo
    
    ' extrae el nombre de la extensión
    Nombre_Archivo = Left(Nombre_Archivo, InStr(Nombre_Archivo, ".") - 1)
    
    ' ruta donde generar el archivo txt con las funciones Api
    Path_txt = App.Path & "\" & Nombre_Archivo & ".txt"
    
    If Path_Dll = vbNullString Or Path_txt = vbNullString Then
        Exit Function
    End If
    
    ' exporta mediante el shell ejecutando el archivo Link.exe
    ret = Shell(Chr(34) & Path_Link & Chr(34) & _
            Comando_Export & Chr(34) & Path_Dll & Chr(34) & _
            " /out:" & Chr(34) & Path_txt & Chr(34), vbHide)
    
    Pausa = Timer
    
    Me.MousePointer = vbHourglass
    ' Pausa para que no dé error al cargar el archivo en el richtextbox
    While (Pausa + 2) > Timer
        DoEvents
    Wend
    ' Carga el txt en el RichtextBox
    RichTextBox1.LoadFile Path_txt
    
    Me.MousePointer = vbNormal
    
    
    MsgBox " Se generó el archivo: " & Nombre_Archivo & ".txt" & vbNewLine & _
           " .. en el App.Path ", vbInformation
           
End Function

Private Sub Form_Load()
    Command1.Caption = " -> Seleccionar archivo para exportar "
End Sub

Private Sub Form_Resize()
    
    Command1.Width = 3000
    Command1.Height = 350
    RichTextBox1.Move 0, 0, _
                      Me.ScaleWidth, _
                      Me.ScaleHeight - (Command1.Height + 100)
    
    Command1.Move Me.ScaleWidth - (Command1.Width + 50), _
                  (RichTextBox1.Top + RichTextBox1.Height + 50)

End Sub

