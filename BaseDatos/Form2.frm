VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Registro"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apellido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   885
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



End Sub

