VERSION 5.00
Begin VB.Form frmbackup 
   BackColor       =   &H8000000A&
   Caption         =   "Backup"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox dirdirec 
      Height          =   990
      Left            =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.DriveListBox drvuni 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3840
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtbyts 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3840
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000E&
      Caption         =   "Option2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Option1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Discket"
      Height          =   195
      Left            =   3840
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Bytes"
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parcial"
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copia"
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Origen"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   465
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

