VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtCadena 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Text            =   "Introduce Texto"
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4335
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   10575
      Begin VB.OptionButton optEspeciales 
         BackColor       =   &H00000000&
         Caption         =   "Caracteres Especiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2760
         Width           =   3135
      End
      Begin VB.OptionButton optAlfabeto 
         BackColor       =   &H00000000&
         Caption         =   "Caracteres Alfa numericos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   4095
      End
      Begin VB.OptionButton optNumeros 
         BackColor       =   &H00000000&
         Caption         =   "Caracteres Numericos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.OptionButton optEspacios 
         BackColor       =   &H00000000&
         Caption         =   "Espacios en Blanco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton optLonguitud 
         BackColor       =   &H00000000&
         Caption         =   "Longuitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Solucion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3495
         Left            =   4440
         TabIndex        =   2
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtResultado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1815
            Left            =   480
            TabIndex        =   12
            Top             =   1440
            Width           =   4575
         End
         Begin VB.Label lblTitulo3 
            BackColor       =   &H00000000&
            Caption         =   "Resultado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   2160
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   9375
      Begin VB.Label lblTitulo2 
         BackColor       =   &H00000000&
         Caption         =   "Introdusca Texto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label lblTitulo1 
      BackColor       =   &H00000000&
      Caption         =   "MANEJO DE CARACTERES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Cadena As String
Public Longitud As Integer
Public Pos As Integer
Private Sub Label2_Click()

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub txtCadena_Click()

txtCadena.Text = Empty


End Sub

Private Sub optLonguitud_Click()
Cadena = txtCadena.Text
Longitud = Len(Cadena)
txtResultado = Str(Longitud) + " Caracteres"
End Sub


