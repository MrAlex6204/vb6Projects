VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Reproductor de Video"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "Ver_Vid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "FF"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "RR"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1200
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   2880
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   327682
      LargeChange     =   1
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Continuar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Detener"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pausar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abrir"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tiempo"
      Height          =   195
      Left            =   5520
      TabIndex        =   6
      Top             =   720
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'*** Constantes ***
Private Const OFN_FILEMUSTEXIST = &H1000&
Private Const OFN_READONLY = &H4&

'*** Variables ***
Private DialogCaption As String
Private FileName As String
Private Const MODAL = 1
Private Const MODELESS = 2

Dim i As Long
Dim ShortName
Dim mssg As String * 255
Dim ResumeStat As String
Dim FFRR As String


   Public Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
      
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

      
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       
       GetShortName = Left(sShortPathName, lRetVal)
   End Function
Private Sub Command1_Click()
  Dim MInfo As String
Screen.MousePointer = 11

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Archivos de Video|*.wmv;*.mpa;*.mpe;*.mpg;*.mpeg;*.avi|Windows Media Video|*.wmv|Archivo de Pelicula(mpeg)|*.mpg;*.mpa;*.mpe;*.mpeg|Video para Windows|*.avi|Todos los ficheros (*.*)|*.*"
CommonDialog1.Flags = OFN_FILEMUSTEXIST Or OFN_READONLY

CommonDialog1.ShowOpen

'#####################################################
 ShortName = GetShortName(CommonDialog1.FileName)
'#####################################################

i = mciSendString("close all", 0&, 0, 0)

Get_Size GetShortName(CommonDialog1.FileName)



Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Screen.MousePointer = 0
Me.Caption = "Reproductor de Video - " + CommonDialog1.FileTitle
App.Title = "Reproductor de Video - " + CommonDialog1.FileTitle
Exit Sub

EH1:

Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
End Sub

Private Sub Command2_Click()
 i = mciSendString("play video1 from " & Slider1.Value, 0&, 0, 0)
End Sub


Private Sub Command3_Click()
 i = mciSendString("pause video1", 0&, 0, 0)
End Sub


Private Sub Command4_Click()
 i = mciSendString("stop video1", 0&, 0, 0)
 i = mciSendString("seek video1 to start", 0&, 0, 0)
 Slider1.Value = 0
End Sub



Private Sub Command5_Click()
 i = mciSendString("resume video1", 0&, 0, 0)
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all off", mssg, 255, 0)
i = mciSendString("status video1 mode", mssg, 255, 0)
FFRR = mssg
Timer2.Enabled = True

End Sub


Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all on", mssg, 255, 0)
Timer2.Enabled = False

Select Case Left$(FFRR, 4)
Case "stop"
 i = mciSendString("stop video1", 0&, 0, 0)
Case "play"
 i = mciSendString("play video1", 0&, 0, 0)
Case "paus"
 i = mciSendString("pause video1", 0&, 0, 0)
Case Else
End Select


End Sub


Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all off", mssg, 255, 0)
i = mciSendString("status video1 mode", mssg, 255, 0)
FFRR = mssg
Timer3.Enabled = True

End Sub


Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = mciSendString("set video1 audio all on", mssg, 255, 0)
Timer3.Enabled = False

Select Case Left$(FFRR, 4)
Case "stop"
 i = mciSendString("stop video1", 0&, 0, 0)
Case "play"
 i = mciSendString("play video1", 0&, 0, 0)
Case "paus"
 i = mciSendString("pause video1", 0&, 0, 0)
Case Else
End Select

End Sub




Private Sub Form_Unload(Cancel As Integer)
 i = mciSendString("close video1", 0&, 0, 0)


End Sub



Public Function Get_Size(ShortName As String)
Dim sReturn As String * 128
Dim lPos As Long
Dim lStart As Long
Dim Last$, Todo$, lWidth, lHeight

    
Last$ = Form1.hWnd & " Style " & &H40000000
Todo$ = "open " & ShortName & " Alias video1 parent " & Last$
i = mciSendString(Todo$, 0&, 0, 0)

i = mciSendString("Where video1 destination", ByVal sReturn, Len(sReturn) - 1, 0)
    
lStart = InStr(1, sReturn, " ")
lPos = InStr(lStart + 1, sReturn, " ")
lStart = InStr(lPos + 1, sReturn, " ")
lWidth = Mid(sReturn, lPos, lStart - lPos)
lHeight = Mid(sReturn, lStart + 1)
    
    
i = mciSendString("put video1 window at 8 80 " & lWidth & " " & lHeight, 0&, 0, 0)


i = mciSendString("set video1 time format ms", 0&, 0, 0)
i = mciSendString("status video1 length", mssg, 255, 0)


Slider1.Max = Val(mssg)

 Timer1.Enabled = True
End Function

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer1.Enabled = False
i = mciSendString("status video1 mode", mssg, 255, 0)

If Left$(mssg, 7) = "playing" Then
  ResumeStat = "playing"
Else
  ResumeStat = ""
End If

i = mciSendString("pause video1", 0&, 0, 0)
End Sub


Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 i = mciSendString("seek video1 to " & Slider1.Value, 0&, 0, 0)

If ResumeStat = "playing" Then
  i = mciSendString("play video1", 0&, 0, 0)
End If
  
   
  Timer1.Enabled = True
   
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Dim VidPos As String
Dim SegunI, MinutosI, SegundosI, LTrackPosition
i = mciSendString("status video1 position", mssg, 255, 0)
VidPos = Str(mssg)

Slider1.Value = VidPos
LTrackPosition = mssg
SegunI = Val(LTrackPosition) \ 1000
MinutosI = SegunI \ 60
SegundosI = SegunI Mod 60
Label1.Caption = MinutosI & " min. " & SegundosI & " seg."
If Err Then Exit Sub

End Sub


Private Sub Timer2_Timer()
On Error Resume Next
 i = mciSendString("stop video1", 0&, 0, 0)
 i = mciSendString("status video1 position", mssg, 255, 0)
 
 If mssg + 50 > Slider1.Max Then
     i = mciSendString("seek video1 to end", 0&, 0, 0)
 Else
    i = mciSendString("play video1 from " & mssg + 50, 0&, 0, 0)
 End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
 i = mciSendString("stop video1", 0&, 0, 0)
 i = mciSendString("status video1 position", mssg, 255, 0)
 
 If mssg - 50 <= 0 Then
    i = mciSendString("seek video1 to start", 0&, 0, 0)
    Slider1.Value = 0
 Else
    i = mciSendString("play video1 from " & mssg - 50, 0&, 0, 0)
 End If

End Sub


