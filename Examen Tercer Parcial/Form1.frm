VERSION 5.00
Begin VB.Form frmArticulos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Captura de Articulos"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12165
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Captura de Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   2520
      TabIndex        =   0
      Top             =   3360
      Width           =   11055
      Begin VB.TextBox txtNumArt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDescrip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox txtPrecio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4200
         TabIndex        =   1
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   7920
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nº de Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   2400
      Picture         =   "Form1.frx":2079
      Top             =   480
      Width           =   11190
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If txtPrecio = Empty Or txtDescrip = Empty Or txtNumArt = Empty Then
MsgBox "Favor de Lenar todos los datos", vbCritical, "VeraSoft Development"

Else
Datos.AbrirBase ("Select * from ListArt")
Datos.BaseDatosOpen.AddNew
Datos.BaseDatosOpen.Fields("Precio") = txtPrecio
Datos.BaseDatosOpen.Fields("Descrip") = txtDescrip
Datos.BaseDatosOpen.Fields("NumArt") = txtNumArt
Datos.BaseDatosOpen.Update
Datos.CerrarBase

txtPrecio = Empty
txtDescrip = Empty
txtNumArt = Empty
End If


End Sub

Private Sub Command2_Click()
txtPrecio = Empty
txtDescrip = Empty
txtNumArt = Empty
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

