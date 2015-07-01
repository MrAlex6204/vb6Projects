VERSION 5.00
Begin VB.Form frmArtVendidos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE DE ARTICULOS"
   ClientHeight    =   5775
   ClientLeft      =   4050
   ClientTop       =   2760
   ClientWidth     =   7680
   Icon            =   "frmArtVendidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7680
   Begin VB.ListBox Listart 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "REPORTE DE ARTICULOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   540
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   5955
   End
End
Attribute VB_Name = "frmArtVendidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
frmArtVendidos.Caption = "REPORTE DE ARTICULOS VENDIDOS DE: " + CajeroOnline
Listart.Clear
ind = 1
Open App.Path + "\CorteCaja.Dat" For Input As #33
            
            Do While Not EOF(33)
     
            Input #33, Index, nArt, Desc, Pre
            Listart.AddItem ind & "<-- CODE:" & nArt & "  " & Desc & " $" & Pre
            ind = ind + 1
            
Loop
Close #33

End Sub
