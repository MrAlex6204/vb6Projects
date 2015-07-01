VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Usb Aplication"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   12825
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":954A
   ScaleHeight     =   8865
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Copy 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9120
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   1080
         TabIndex        =   44
         Top             =   0
         Width           =   120
      End
      Begin VB.Label CopyFile 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Copiar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   240
         TabIndex        =   43
         Top             =   120
         Width           =   750
      End
   End
   Begin VB.FileListBox filelist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      ForeColor       =   &H00C0C000&
      Height          =   5685
      Hidden          =   -1  'True
      Left            =   9480
      Pattern         =   "*.mpg;*.mpeg;*.mp4;*.wav;*.mpeg;*.wmv;*.wma;*.mp3;*.wav;*.aif;*.au;*.3gp;*.avi;*.flv;*.fvl"
      System          =   -1  'True
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame MediaControls 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   1200
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   7335
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   1080
         TabIndex        =   37
         ToolTipText     =   "Volumen"
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         OLEDropMode     =   1
         Max             =   100
      End
      Begin VB.Label Play 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   2520
         TabIndex        =   35
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Abrir 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   1200
         TabIndex        =   34
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Pause 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   3720
         TabIndex        =   33
         Top             =   360
         Width           =   585
      End
      Begin VB.Label RepStop 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   5160
         TabIndex        =   32
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame frmControls 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1200
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Escritorio"
         DragIcon        =   "Form1.frx":1AF2FC
         ForeColor       =   &H00808000&
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Estirado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Left            =   4320
            TabIndex        =   26
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Mozaico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Left            =   3000
            TabIndex        =   25
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Restablecer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Centrado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.DriveListBox drvDiks 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3240
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   5160
         TabIndex        =   30
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblPath 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione la Unidad  para aplicar el fondo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el Texto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   2760
         TabIndex        =   27
         Top             =   360
         Width           =   1785
      End
   End
   Begin VB.Frame MenuAplicar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   4560
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      Begin VB.Label LabelIcono 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Icono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Width           =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808000&
         X1              =   0
         X2              =   1920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         X1              =   0
         X2              =   1920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Escritorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Acerca 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   10920
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "VeraSoft Development"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "USB Aplication"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   435
         Left            =   2160
         TabIndex        =   18
         Top             =   1080
         Width           =   2760
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "USB Aplication Solo Corre con Windows XP ó Windows 98 y  Windows 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   795
         Left            =   2160
         TabIndex        =   17
         Top             =   1800
         Width           =   2760
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Desarrollado Por: Oscar Alejandro Vera Hdz."
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   3600
         Width           =   3240
      End
   End
   Begin VB.Frame MenuAyuda 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5760
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Acerca de USB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1650
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2055
      Left            =   13680
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   9120
      Top             =   7320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9120
      Top             =   6840
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   13200
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   """Imagenes (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif"""
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   1440
      TabIndex        =   38
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   8760
      TabIndex        =   36
      Top             =   6840
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Shape Controles 
      BackColor       =   &H00000000&
      BorderColor     =   &H00808000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   1080
      Top             =   6720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   1680
      TabIndex        =   13
      Top             =   1320
      Width           =   930
   End
   Begin VB.Image PIC 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4215
      Left            =   1440
      Picture         =   "Form1.frx":1B0AAE
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   7455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   7800
      TabIndex        =   7
      Top             =   480
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Reproductor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Examinar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   12960
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

      
' Constantes para los flags para reproducir sonido
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  look for application specific association
Private Const SND_APPLICATION = &H80
'  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS = &H10000
'  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_ID = &H110000
'  play asynchronously
Private Const SND_ASYNC = &H1
  '  play synchronously (default)
Private Const SND_SYNC = &H0

'  name is a file name
Private Const SND_FILENAME = &H20000
'  loop the sound until next sndPlaySound
Private Const SND_LOOP = &H8
'  lpszSoundName points to a memory file
Private Const SND_MEMORY = &H4
'  silence not default, if sound not found
Private Const SND_NODEFAULT = &H2
 '  don't stop any currently playing sound
Private Const SND_NOSTOP = &H10
 '  don't wait if the driver is busy
Private Const SND_NOWAIT = &H2000
 '  purge non-static events for task
Private Const SND_PURGE = &H40
 '  name is a resource name or atom
Private Const SND_RESOURCE = &H40004

' Declaración del api PlaySound
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Reproduce, RepVisible As Boolean
' Reproduce el archivo de sonido wav
Sub Reproducir_WAV(Archivo As String, flags As Long)
    
    Dim ret As Long
    ' Le pasa el path y los flags al api
    ret = PlaySound(Archivo, ByVal 0&, flags)
End Sub
Sub RepSound(File As String)
If Reproduce = True Then

Else

Call Reproducir_WAV(File, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT)

End If

End Sub



 Sub Aceptar_Click()

End Sub

Private Sub Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

AceptarBotton = True
Aceptar.Picture = LoadPicture(App.Path + "\api\Aceptar.bmp")
End Sub

 Sub Aplicar_Click()

 
End Sub

Private Sub Aplicar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
AplicarBotton = True
Aplicar.Picture = LoadPicture(App.Path + "\api\Aplicar.bmp")
End Sub

Private Sub Ayuda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
AyudaBotton = True
Ayuda.Picture = LoadPicture(App.Path + "\api\Ayuda.bmp")
End Sub

 Sub Cerrar_Click()

End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
CerrarBotton = True
Cerrar.Picture = LoadPicture(App.Path + "\api\cerrar.bmp")
End Sub





 Sub Examinar_Click()

End Sub

Private Sub Examinar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
ExaminarBotton = True
Examinar.Picture = LoadPicture(App.Path + "\api\Examinar.bmp")

End Sub


Private Sub Abrir_Click()
Form2.Abrir

If Form2.RepDialogo.FileName = "" Then
MsgBox "Seleccione algun archivo", vbCritical, "USB aplication"
Else
filelist.Path = Form2.RepDialogo.DefaultExt
filelist.Visible = True
lblStatus.Visible = True
lblStatus = Form2.Reproductor.Status
Form2.PLAYFILE
End If

End Sub

Private Sub Abrir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Abrir.FontSize = 12
Abrir.FontBold = True
Abrir.ForeColor = &HC0C000
End Sub

Private Sub Acerca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuAyuda.Visible = False
MenuAplicar.Visible = False
Reproduce = False
 Label20.FontSize = 10
Label20.FontBold = False
Label20.ForeColor = &H808000

 Label21.FontSize = 10
Label21.FontBold = False
Label21.ForeColor = &H808000

 Label22.FontSize = 10
Label22.FontBold = False
Label22.ForeColor = &H808000


End Sub



Private Sub CopFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

CopFile.FontSize = 12
CopFile.FontBold = True
CopFile.ForeColor = &HC0C000
End Sub

Private Sub Copy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CopyFile.FontSize = 10
CopyFile.FontBold = False
CopyFile.ForeColor = &H808000
End Sub

Private Sub CopyFile_Click()
Shell "explorer.exe " + filelist.Path + "\" + filelist.FileName, vbNormalFocus
Copy.Visible = False
End Sub

Private Sub CopyFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

CopyFile.FontSize = 12
CopyFile.FontBold = True
CopyFile.ForeColor = &HC0C000
End Sub

Private Sub filelist_Click()
Form2.Reproductor.URL = filelist.Path + "\" + filelist.FileName
End Sub

Private Sub filelist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Copy.Visible = True
Copy.Left = filelist.Left + X + 5
Copy.Top = filelist.Top + Y + 1
End If
End Sub

Private Sub Form_Activate()

If RepVisible = True Then
Form2.Show
Else
Unload Form2
End If
End Sub

Private Sub Form_Load()
On Error GoTo ERR
RepVisible = False
'Color magneta del form &H00FF00FF&
'Me.Width = 10095
'Me.Height = 6855



    
   MakeFormTransparent Me, vbMagenta
   
   
 

PIC.Appearance = 0
lblPath = Empty
lblNombre = Empty
Exit Sub
ERR:
MsgBox "Cierre y Vuelta a Intentarlo", vbInformation, "VeraSoftDevelopment"
Shell App.Path + "\Config.Bat", vbHide



End Sub

 Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MenuAyuda.Visible = False
 MenuAplicar.Visible = False
Reproduce = False

Frame1.FontSize = 10
Frame1.FontBold = False
Frame1.ForeColor = &H808000

 Label1.FontSize = 10
Label1.FontBold = False
Label1.ForeColor = &H808000
 
 Label2.FontSize = 10
Label2.FontBold = False
Label2.ForeColor = &H808000

Label2.FontSize = 10
Label2.FontBold = False
Label1.ForeColor = &H808000

Label3.FontSize = 10
Label3.FontBold = False
Label3.ForeColor = &H808000

Label4.FontSize = 10
Label4.FontBold = False
Label4.ForeColor = &H808000

Label5.FontSize = 10
Label5.FontBold = False
Label5.ForeColor = &H808000

Label6.FontSize = 10
Label6.FontBold = False
Label6.ForeColor = &HC0C000
 
 Label11.FontSize = 10
Label11.FontBold = False
Label11.ForeColor = &HC0C000

 Label12.FontSize = 10
Label12.FontBold = False
Label12.ForeColor = &H808000
 
  Label14.FontSize = 10
Label14.FontBold = False
Label14.ForeColor = &H808000

 Label15.FontSize = 10
Label15.FontBold = False
Label15.ForeColor = &H808000

 Label16.FontSize = 10
Label16.FontBold = False
Label16.ForeColor = &H808000

 Label21.FontSize = 10
Label21.FontBold = False
Label21.ForeColor = &H808000

 Label22.FontSize = 10
Label22.FontBold = False
Label22.ForeColor = &H808000
Reproduce = False
 
Dim lngReturnValue As Long


        If Button = 1 Then
        
      Mover.MoverForm Form1.hWnd
     
        End If
        



 
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Reproduce = False
Frame1.FontSize = 12
Frame1.FontBold = True
Frame1.ForeColor = &HC0C000

Label14.FontSize = 10
Label14.FontBold = False
Label14.ForeColor = &H808000

 Label15.FontSize = 10
Label15.FontBold = False
Label15.ForeColor = &H808000

 Label16.FontSize = 10
Label16.FontBold = False
Label16.ForeColor = &H808000

 Label23.FontSize = 10
Label23.FontBold = False
Label23.ForeColor = &H808000

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CopFile.FontSize = 10
CopFile.FontBold = False
CopFile.ForeColor = &H808000
End Sub

Private Sub frmControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Reproduce = False
 Label11.FontSize = 10
Label11.FontBold = False
Label11.ForeColor = &HC0C000

Label24.FontSize = 10
Label24.FontBold = False
Label24.ForeColor = &HC0C000

Label10.FontSize = 10
Label10.FontBold = False
Label10.ForeColor = &HC0C000

End Sub



Private Sub Label1_Click()


dialogo.ShowOpen

PIC.Picture = LoadPicture(dialogo.FileName)
lblPath.Visible = True
lblPath = dialogo.FileName
lblNombre = dialogo.FileTitle



End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Label1.FontSize = 12
Label1.FontBold = True
Label1.ForeColor = &HC0C000
End Sub

Private Sub Label10_Click()

    'La variable El_Color almacenará el color en formato Long
    'del color elegido. Si no se eligió ninguno retornamos desde
    'la función el valor -1, si no establecemos el color defondo
    'del form pasandole el valor devuelto por la función
    
    ' llamamos al cuadro diálogo Seleccionar Color
    
    'El_Color es una variable publica declarada en el modulo
    
    El_Color = Abrir_CommonDialog_Color(Me)
    
    If El_Color <> -1 Then
    
        ' establecemos el color de fondo del Form con el color seleccionado
        
        lblColor.BackColor = El_Color
    Else
        MsgBox "Se canceló ", vbInformation, "USB Aplication"
    End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label10.FontSize = 12
Label10.FontBold = True
Label10.ForeColor = &HFFFF80
End Sub

Private Sub Label11_Click()
frmControls.Visible = False
Controles.Visible = False
Frame1.Visible = False
MediaControls.Visible = False
lblPath.Visible = False
Label11.Visible = False
RepVisible = False
lblNombre = Empty
Unload Form2
lblStatus.Visible = False
filelist.Visible = False
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label11.FontSize = 12
Label11.FontBold = True
Label11.ForeColor = &HFFFF80
End Sub

Private Sub Label12_Click()
Acerca.Visible = True
Acerca.Width = 6975
Acerca.Height = 3855
Acerca.Left = 1560
Acerca.Top = 1800
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Label12.FontSize = 12
Label12.FontBold = True
Label12.ForeColor = &HC0C000

End Sub

Private Sub Label14_Click()
'Para sacar el papel Tapiz se le envía una cadena vacía en lpvParam
  SystemParametersInfo SPI_SETDESKWALLPAPER, 0, "", _
  SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label14.FontSize = 12
Label14.FontBold = True
Label14.ForeColor = &HC0C000
End Sub

Private Sub Label15_Click()
If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
Escritorio.cambiarTapiz dialogo.FileName, 0
End If
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label15.FontSize = 12
Label15.FontBold = True
Label15.ForeColor = &HC0C000
End Sub



Private Sub Label16_Click()
If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
Escritorio.cambiarTapiz dialogo.FileName, 1
End If
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label16.FontSize = 12
Label16.FontBold = True
Label16.ForeColor = &HC0C000
End Sub

 Sub Label2_Click()
Form2.Show
Label11.Visible = True
MediaControls.Visible = True
Controles.Visible = True
Label1.Visible = True
Slider1.Value = Form2.Reproductor.settings.volume
RepVisible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label2.FontSize = 12
Label2.FontBold = True
Label2.ForeColor = &HC0C000
End Sub

Private Sub Label20_Click()
Acerca.Visible = False
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label20.FontSize = 12
Label20.FontBold = True
Label20.ForeColor = &HFFFF80
End Sub

Private Sub Label21_Click()
Label11.Visible = True
lblPath.Visible = True

frmControls.Visible = True
Frame1.Visible = False
Controles.Visible = True




End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Label21.FontSize = 12
Label21.FontBold = True
Label21.ForeColor = &HC0C000
End Sub

Private Sub Label22_Click()
lblPath.Visible = True
Label11.Visible = True
frmControls.Visible = True
Controles.Visible = True
Frame1.Visible = True


If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
Frame1.Visible = True
End If
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Label22.FontSize = 12
Label22.FontBold = True
Label22.ForeColor = &HC0C000
End Sub

Private Sub Label23_Click()
If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
Escritorio.cambiarTapiz dialogo.FileName, 2
End If
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label23.FontSize = 12
Label23.FontBold = True
Label23.ForeColor = &HC0C000
End Sub

Private Sub Label24_Click()
On Error Resume Next
If Form1.lblPath = Empty Then
MsgBox "Porfavor Seleccione una Imagen", vbCritical, "Seleccione una Imagen"
Else
'
FileCopy Form1.lblPath.Caption, drvDiks.Drive + "\" + dialogo.FileTitle

Open drvDiks.Drive + "\DESKTOP.INI" For Output As #1 'genera el archivo el el drive
'seleccionado por drvDisk
Dim Color As String
Color = Str(El_Color)
Print #1, "[{BE098140-A513-11D0-A3A4-00C04FD706EC}]"
Print #1, "ICONAREA_IMAGE=" + dialogo.FileTitle
Print #1, "ICONAREA_TEXT=" + Color

Close #1
MsgBox "Fondo aplicado Porfavor Actualiza la Unidad Para Ver los Cambios", vbExclamation_, "Fondo Aplicado"
Shell "explorer.exe " + drvDiks.Drive, vbMaximizedFocus
End If
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label24.FontSize = 12
Label24.FontBold = True
Label24.ForeColor = &HFFFF80
End Sub

Private Sub Label25_Click()
Form2.Reproductor.Controls.stop
End Sub

Private Sub Label26_Click()
Copy.Visible = False
End Sub



Private Sub Label3_Click()
MenuAplicar.Visible = True
        
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuAplicar.Visible = True
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label3.FontSize = 12
Label3.FontBold = True
Label3.ForeColor = &HC0C000
End Sub

Private Sub Label4_Click()
MenuAyuda.Visible = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuAyuda.Visible = True
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label4.FontSize = 12
Label4.FontBold = True
Label4.ForeColor = &HC0C000
End Sub

Private Sub Label5_Click()
Form1.WindowState = 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label5.FontSize = 12
Label5.FontBold = True
Label5.ForeColor = &HFFFF80
End Sub

Private Sub Label6_Click()
Unload Form2
Unload Form3
End
End Sub

Sub Min_Click()



 
End Sub

Private Sub Min_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound1.wav"
Reproduce = True
MinimizarBotton = True
Min.Picture = LoadPicture(App.Path + "\api\Minimizar.bmp")
End Sub



Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True
Label6.FontSize = 12
Label6.FontBold = True
Label6.ForeColor = &HFFFF80
End Sub

Private Sub Picture1_Click()

End Sub



Private Sub lblNameFile_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub LabelIcono_Click()
Form3.Show
End Sub

Private Sub LabelIcono_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

LabelIcono.FontSize = 12
LabelIcono.FontBold = True
LabelIcono.ForeColor = &HC0C000
End Sub

Private Sub MediaControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Reproduce = False
Abrir.FontSize = 10
Abrir.FontBold = False
Abrir.ForeColor = &H808000

Play.FontSize = 10
Play.FontBold = False
Play.ForeColor = &H808000

Pause.FontSize = 10
Pause.FontBold = False
Pause.ForeColor = &H808000

RepStop.FontSize = 10
RepStop.FontBold = False
RepStop.ForeColor = &H808000
End Sub

Private Sub MenuAplicar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label21.FontSize = 10
Label21.FontBold = False
Label21.ForeColor = &H808000

 Label22.FontSize = 10
Label22.FontBold = False
Label22.ForeColor = &H808000
Reproduce = False

LabelIcono.FontSize = 10
LabelIcono.FontBold = False
LabelIcono.ForeColor = &H808000

End Sub

Private Sub MenuAyuda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label12.FontSize = 10
Label12.FontBold = False
Label12.ForeColor = &H808000
Reproduce = False
End Sub

Private Sub NombreMedia_Click()

End Sub

Private Sub Pause_Click()
Form2.PAUSEFILE
End Sub

Private Sub Pause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Pause.FontSize = 12
Pause.FontBold = True
Pause.ForeColor = &HC0C000
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MenuAyuda.Visible = False
 MenuAplicar.Visible = False
End Sub

Private Sub Play_Click()
Form2.PLAYFILE

End Sub

Private Sub Play_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

Play.FontSize = 12
Play.FontBold = True
Play.ForeColor = &HC0C000
End Sub

Private Sub Stop_Click()

End Sub

Private Sub Stop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub RepStop_Click()
Form2.STOPFILE
End Sub

Private Sub RepStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RepSound App.Path + "\Sounds\Sound2.wav"
Reproduce = True

RepStop.FontSize = 12
RepStop.FontBold = True
RepStop.ForeColor = &HC0C000
End Sub

Private Sub Slider1_Scroll()
Form2.VOLUMEN (Slider1.Value)
End Sub

Private Sub SmartMenuXP1_Click(ByVal ID As Long)

End Sub

Private Sub Timer1_Timer()
'Call Reproducir_WAV(App.Path + "\Sounds\Sound3.wav", SND_FILENAME Or SND_NOSTOP Or SND_NODEFAULT)

Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
   Form2.Top = Form1.Top + 1690
   
   Form2.Left = Form1.Left - Form2.Width + 9000
End Sub

Private Sub VScroll1_Change()
 Form2.Top = Form1.Top + VScroll1.Value
 Label7 = VScroll1.Value
   
   Form2.Left = Form1.Left
End Sub
