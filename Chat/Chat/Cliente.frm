VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   Caption         =   "Cliente"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   Icon            =   "Cliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   9720
      Top             =   4920
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   9120
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "Cliente.frx":0CCA
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Marco 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   1
      Left            =   10200
      ScaleHeight     =   3375
      ScaleWidth      =   2655
      TabIndex        =   22
      ToolTipText     =   "Clic para cambiar la imagen"
      Top             =   4920
      Width           =   2655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   240
         ScaleHeight     =   135
         ScaleWidth      =   1515
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   1515
         Begin VB.Label LabelProgres 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   33
            Top             =   -30
            Width           =   1065
         End
         Begin VB.Label LabelProgres 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   34
            Top             =   -30
            Width           =   1335
         End
         Begin VB.Label LabelProgres 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   35
            Top             =   -30
            Width           =   1605
         End
      End
   End
   Begin VB.PictureBox Marco 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   0
      Left            =   10200
      ScaleHeight     =   3375
      ScaleWidth      =   2655
      TabIndex        =   21
      Top             =   1080
      Width           =   2655
      Begin VB.PictureBox picVolumen 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2200
         Picture         =   "Cliente.frx":0ECE
         ScaleHeight     =   255
         ScaleWidth      =   150
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   1300
         X2              =   2400
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.PictureBox Bootons 
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   4
      Left            =   5280
      Picture         =   "Cliente.frx":1043
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   20
      ToolTipText     =   "Opciones"
      Top             =   120
      Width           =   825
   End
   Begin VB.PictureBox Bootons 
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   3
      Left            =   4200
      Picture         =   "Cliente.frx":1892
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   19
      ToolTipText     =   "Enviar Adjunto"
      Top             =   120
      Width           =   825
   End
   Begin VB.PictureBox Bootons 
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   2
      Left            =   3120
      Picture         =   "Cliente.frx":206E
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   18
      ToolTipText     =   "Iniciar Audio"
      Top             =   120
      Width           =   825
   End
   Begin VB.PictureBox Bootons 
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   1
      Left            =   2040
      Picture         =   "Cliente.frx":28F7
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   17
      ToolTipText     =   "Iniciar WebCam"
      Top             =   120
      Width           =   825
   End
   Begin VB.PictureBox Bootons 
      BorderStyle     =   0  'None
      Height          =   825
      Index           =   0
      Left            =   960
      Picture         =   "Cliente.frx":3168
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   16
      ToolTipText     =   "Desconectar"
      Top             =   120
      Width           =   825
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H8000000F&
      Height          =   4095
      Left            =   3840
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Utilizar Emoticons"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   3600
         Picture         =   "Cliente.frx":3979
         ToolTipText     =   "Agregar nuevos"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   2520
         Picture         =   "Cliente.frx":437B
         ToolTipText     =   "Buscar"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Emoticons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   3120
         Picture         =   "Cliente.frx":4D7D
         ToolTipText     =   "Eliminar"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "Cliente.frx":577F
         ToolTipText     =   "Insertar"
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   3480
         Picture         =   "Cliente.frx":6181
         ToolTipText     =   "Cerrar Cuador"
         Top             =   120
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bold"
            Object.ToolTipText     =   "Negrita"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "underline"
            Object.ToolTipText     =   "Subrayado"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Desacer"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imagen"
            Object.ToolTipText     =   "Insertar imagen"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Objeto"
            Object.ToolTipText     =   "Insertar objeto"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Caritas"
            Object.ToolTipText     =   "Emoticons"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sumbido"
            Object.ToolTipText     =   "Enviar un sumbido"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageCombo ImageCombo3 
         Height          =   330
         Left            =   8280
         TabIndex        =   7
         ToolTipText     =   "Color"
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   6
         ToolTipText     =   "Tamaño"
         Top             =   0
         Width           =   855
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "Fuente"
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   3120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3600
      Top             =   720
   End
   Begin VB.PictureBox Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   8160
      ScaleHeight     =   1095
      ScaleWidth      =   1455
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
      Begin VB.PictureBox PictureEnviar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   360
         ScaleHeight     =   885
         ScaleWidth      =   990
         TabIndex        =   30
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2280
      Top             =   720
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1800
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   8400
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   7800
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7200
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":6B83
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":711D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":76B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":7C51
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":81EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":8785
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":8D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":92B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":9853
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":9DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Cliente.frx":A187
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   5160
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   555
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   2778
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Cliente.frx":A721
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7435
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Cliente.frx":A7A3
      MouseIcon       =   "Cliente.frx":A825
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picSplit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   0
      MouseIcon       =   "Cliente.frx":A987
      MousePointer    =   99  'Custom
      ScaleHeight     =   210
      ScaleWidth      =   9885
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5640
      Width           =   9885
   End
   Begin VB.PictureBox Borde 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   9495
      TabIndex        =   13
      Top             =   1120
      Width           =   9495
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   90
      End
      Begin VB.Image Comando 
         Height          =   210
         Index           =   3
         Left            =   9120
         Picture         =   "Cliente.frx":AAD9
         ToolTipText     =   "Expandir"
         Top             =   60
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   300
         Index           =   0
         Left            =   145
         TabIndex        =   25
         Top             =   30
         Width           =   90
      End
   End
   Begin VB.PictureBox Borde 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   1
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   9495
      TabIndex        =   14
      Top             =   5880
      Width           =   9495
   End
   Begin VB.PictureBox Borde 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   2
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   9495
      TabIndex        =   15
      Top             =   7920
      Width           =   9495
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   135
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   165
         Width           =   75
      End
   End
   Begin VB.PictureBox Foto 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2300
      Index           =   0
      Left            =   10320
      Picture         =   "Cliente.frx":AD49
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   23
      Top             =   1560
      Width           =   2300
   End
   Begin VB.PictureBox Foto 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2300
      Index           =   1
      Left            =   10440
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   24
      Top             =   5520
      Width           =   2300
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   4200
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Comando 
      Height          =   210
      Index           =   0
      Left            =   11880
      Picture         =   "Cliente.frx":EC6A
      ToolTipText     =   "Minimizar"
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Comando 
      Height          =   210
      Index           =   1
      Left            =   12240
      Picture         =   "Cliente.frx":EEE7
      ToolTipText     =   "Maximizar"
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Comando 
      Height          =   210
      Index           =   2
      Left            =   12600
      Picture         =   "Cliente.frx":F2F3
      ToolTipText     =   "Cerrar"
      Top             =   240
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Comando 
      Height          =   225
      Index           =   4
      Left            =   120
      Picture         =   "Cliente.frx":F6FB
      ToolTipText     =   "Ocultar el marco de la ventana"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Boton 
      Height          =   885
      Index           =   0
      Left            =   8640
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image Boton 
      Height          =   900
      Index           =   1
      Left            =   7200
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image Boton 
      Height          =   900
      Index           =   2
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DX As DirectX8
Private DS As DirectSound8
Private DSToneBuffer As DirectSoundSecondaryBuffer8
Private desc As DSBUFFERDESC


'----------volumen
Private Declare Function waveOutSetVolume Lib "winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Private Declare Function waveOutGetVolume Lib "winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
'------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Ejecutar
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'----va con sendmesage y es para pegar el portapapeles
Private Const WM_PASTE = &H302
'-------redondear
Private Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long, _
ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
(ByVal hWnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long
'---------mover el form en forma de mascara
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'---------------------soltar mouse
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal DX As Long, ByVal dy As Long, _
ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Const MOUSEEVENTF_LEFTUP = &H4

'----------parpadeo de la ventana
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Const Invert = 1

'----------extraerIcono
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Private Const SHGFI_LARGEICON = &H0           ' get large icon
Private Const SHGFI_SMALLICON = &H1           ' get small icon
Private Const SHGFI_SYSICONINDEX = &H4000        ' get system icondex
Private Const ILD_TRANSPARENT = &H1
Private Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long           ' : icon
    iIcon As Long     ' : icondex
    dwAttributes As Long        ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80     ' : type name
End Type
'------------------------------------
Dim strRutaImagen As String
Dim ancho As Single, alto As Single, porcentaje As Single
Dim Fotito As IPictureDisp
Dim Toolsbar As IPictureDisp
Dim MiNick As String

Dim rtf As String
Dim Ruta As String
Dim Imagen() As Byte
Dim RutaArchivo As String
Dim NombreUsuario As String
Dim User As Boolean

Dim RutaAudio As String
Dim AudioB() As Byte
Dim Archivo() As Byte
Dim RutaSonido As String
Dim FlagAudio As Boolean
Dim UnaVes As Boolean
Dim Mascara As Boolean
Dim Escriviendo As Boolean
Dim WebCam As Boolean
Dim Audio As Boolean
Dim vol As String

Dim Adjunto As Boolean
Dim NombreArchivo As String
Dim TamañoArchivo As Long
Dim NombreAdjunto As String
Dim TamañoAdjunto As Long
Dim NombreAdj As Boolean
Dim TamañoAdj As Boolean
Dim nn As String
Dim vv As String
Dim RutaGuardar As String
Dim Progreso As Long

Dim i As Integer
Dim Text As String
Dim StrInfo As String

Dim nn2 As String, vv2 As String
Dim Progreso2 As Long
Dim TamañoAdjunto2 As Long



Private Sub Redondear(Shape As PictureBox)
Dim lRet As Long
Dim l As Long
Dim Width As Long
Dim Height As Long

Width = Shape.Width / Screen.TwipsPerPixelX
Height = Shape.Height / Screen.TwipsPerPixelY

lRet = CreateRoundRectRgn(0, 0, Width, Height, 14, 14)
l = SetWindowRgn(Shape.hWnd, lRet, True)

Dim bgdImage As Picture
Dim X           As Integer
Dim y           As Integer


Set bgdImage = Shape.Picture
y = 0
While y < Shape.Height
    X = 0
    While X < Shape.Width
        Shape.PaintPicture bgdImage, X, y
        X = X + bgdImage.Width \ 2
    Wend
        y = y + bgdImage.Height \ 2
Wend

End Sub


Private Sub Bootons_Click(Index As Integer)
On Error GoTo salir
Select Case Index
Case 0
Winsock1.Close
Winsock2.Close
Winsock3.Close
Timer1.Enabled = False
Timer2.Enabled = False
Call CloseDevice
Picture2.Visible = False
Case 1
If WebCam = False Then
    WebCam = True
    Enviar ("#WebCam#")
    TextInfo "Esperando que " & NombreUsuario & "acepte ver tu WebCam..."
    
Else
    WebCam = False
    Timer1.Enabled = False
    DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
    Close #1
    TextInfo "As desconectado tu WebCam."
    MostrarFoto
   
End If
Case 2
If Not FlagAudio Then
FlagAudio = True
Enviar ("#Audio#")
TextInfo "Esperando que " & NombreUsuario & "acepte oir tu vos..."
Else
FlagAudio = False
Timer2.Enabled = False
RECORD_Save
Enviar ("#FinAudio#")
End If

Case 3
If Adjunto = True Then TextInfo "Solo puede enviar un adjunto a la ves.": Exit Sub
Dim Extenciones As String
Extenciones = "Todos los archivos" + Chr$(0) + "*.*"
RutaArchivo = ShowOpen(Extenciones)
NombreArchivo = Right(RutaArchivo, Len(RutaArchivo) - InStrRev(RutaArchivo, "\"))
TamañoArchivo = FileLen(RutaArchivo)
Enviar ("#Adjunto#")
RichTextBox1.SelStart = Len(RichTextBox1.Text)
TextInfo ("Esperando que " & NombreUsuario & " acepte.")
'-------
PegarIcono RutaArchivo
TextOkCancel NombreArchivo, "(" & Trim(Format$(Format$((TamañoArchivo \ 1024) + 1, "##,###,##0") & " KB)", "@@@@@@@@@@@@"))
RichTextBox1.SelRTF = Text1
TextOkCancel "Cancelar_", ""
Adjunto = True
Case 4
PopupMenu Form2.mnuMascaras, , Bootons(4).Left, Bootons(4).Top + 800
End Select
Exit Sub
salir:

End Sub



Private Sub Bootons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Bootons(Index).Left = Bootons(Index).Left + 30
Boton(2).Picture = Boton(1).Picture
End Sub

Private Sub Bootons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If UnaVes Then
Boton(2).Left = Bootons(Index).Left - 130
Boton(2).Picture = Boton(0).Picture
Boton(2).Visible = True
UnaVes = False
End If
End Sub

Private Sub Bootons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Bootons(Index).Left = Bootons(Index).Left - 30
Boton(2).Picture = Boton(0).Picture
End Sub


Private Sub Comando_Click(Index As Integer)
Select Case Index
Case 0
Me.WindowState = 1
Case 1
Me.WindowState = 2
Comando_Click (4)
Case 2
Unload Me
Case 3


If Marco(0).Visible Then
        For i = 0 To 1
            Marco(i).Visible = False
            Foto(i).Visible = False
        Next
    Else
        For i = 0 To 1
            Marco(i).Visible = True
            Foto(i).Visible = True
        Next
    End If
    Form_Resize

Case 4
    Dim lRet As Long
    Dim l As Long
    Dim Width As Long
    Dim Height As Long
    Width = Me.Width / Screen.TwipsPerPixelX
    Height = Me.Height / Screen.TwipsPerPixelY
    If Not Mascara Then
        lRet = CreateRoundRectRgn(5, 30, Width - 7, Height - 7, 30, 30)
        l = SetWindowRgn(Me.hWnd, lRet, True)
        For l = 0 To 2
            Comando(l).Visible = True
        Next
        Mascara = True
    Else
        l = SetWindowRgn(Me.hWnd, 0, False)
        For l = 0 To 2
            Comando(l).Visible = False
        Next
        Mascara = False
    End If
End Select
End Sub

Private Sub Send()
If RichTextBox2.Text = "" Then Beep: Exit Sub
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox2.SelStart = 0
RichTextBox2.SelRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fmodern\fprq1\fcharset0 Lucida Console;}{\f1\fnil\fcharset0 MS Sans Serif;}}" & vbCrLf & _
"{\colortbl ;\red0\green255\blue0;\red255\green0\blue255;}" & vbCrLf & _
"\viewkind4\uc1\pard\cf1\highlight2\lang3082\b\f0\fs28 " & MiNick & ": \highlight0" & vbCrLf & "\cf0\b0\f1\fs17" & vbCrLf & "\par }"

Enviar (RichTextBox2)
RichTextBox1.SelRTF = RichTextBox2
RichTextBox1.SelStart = Len(RichTextBox1.Text)
Enviar ("ter3")
RichTextBox2 = ""
RecuperarInfo
Escriviendo = False

End Sub



Private Sub InsertarImagen()
On Error GoTo ErrorImagen
Dim strContenidoPortapapeles As String
Picture1.Width = 3200: Picture1.Height = 3200
CargarImagen Picture1

ancho = (ancho * porcentaje) / 100
alto = (alto * porcentaje) / 100
Picture1.Width = ancho: Picture1.Height = alto
Picture1.PaintPicture Fotito, 0, 0, ancho, alto
strContenidoPortapapeles = Clipboard.GetText
If WebCam Then Timer1.Enabled = False
Clipboard.Clear
Clipboard.SetData Picture1.Image
SendMessage RichTextBox2.hWnd, WM_PASTE, 0, 0
Clipboard.Clear
Clipboard.SetText strContenidoPortapapeles
If WebCam Then Timer1.Enabled = True
Exit Sub
ErrorImagen:
If err.Number <> 91 Then
    MsgBox "Error " & err.Number & " " & err.Description
    Exit Sub
End If

End Sub





Private Sub TextOkCancel(Cadena1 As String, Cadena2 As String)
RichTextBox1.SelStart = Len(RichTextBox1.Text)
With RichTextBox1
.SelBold = True
.SelColor = &HFF0000
.SelUnderline = True
.SelFontSize = 12
.SelText = Cadena1
.SelUnderline = False
.SelText = Space(10)
.SelUnderline = True
.SelText = Cadena2 & vbCrLf & vbCrLf
End With
End Sub



Private Sub TextInfo(Cadena1 As String, Optional Index As Integer)
RichTextBox1.SelStart = Len(RichTextBox1.Text)
If Index > 0 Then
Clipboard.Clear
Clipboard.SetData ImageList4.ListImages(Index)
RichTextBox1.Locked = False
SendMessage RichTextBox1.hWnd, WM_PASTE, 0, 0
RichTextBox1.Locked = True
End If
With RichTextBox1
.SelBold = True
.SelColor = 13209
.SelFontSize = 12
.SelText = "__________" & vbCrLf & Cadena1 & vbCrLf & vbCrLf

End With
End Sub










Private Sub Form_Load()
On Error Resume Next
Form2.Text1(0) = Form1.Winsock1.LocalIP
Set DX = New DirectX8

Set DS = DX.DirectSoundCreate("")


DS.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL


desc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_GLOBALFOCUS






For i = 0 To 4
Call CalcPic(Bootons(i))
Next
Call CalcPic(picVolumen)

Foto(1) = Foto(0)
CargarEmoticons
cargarImagenes (0)
ComboFuentes
ComboColores
ComboTamaño
Call InitSigns
RichTextBox2.SelFontSize = (GetSetting(App.EXEName, "Tamaño", "TamañoFuente"))
RichTextBox2.SelColor = Val((GetSetting(App.EXEName, "Color", "ColorFuente")))
RichTextBox2.SelFontName = (GetSetting(App.EXEName, "Fuente", "NombreFuente"))
UpdateTextInfo

'--------------------
RutaAudio = App.Path & "\temporal3.Wav"
Ruta = App.Path & "\temporal2.bmp"
End Sub

Sub cargarImagenes(Index As Integer)
Set Toolsbar = LoadPicture(App.Path & "\Barras\" & Index & ".gif")
Me.Picture = LoadPicture(App.Path & "\Fondos\" & Index & ".gif")
Borde(0).Picture = LoadPicture(App.Path & "\Bordes\" & Index & ".gif")
Borde(1).Picture = Borde(0).Picture
Borde(2).Picture = Borde(0).Picture
Form_Resize
Marco(0).Picture = LoadPicture(App.Path & "\Marcos\" & Index & ".gif")
Marco(1).Picture = Marco(0).Picture
Call CalcPic(Marco(0))
Call CalcPic(Marco(1))
Boton(0).Picture = LoadPicture(App.Path & "\Botones\R" & Index & ".gif")
Boton(1).Picture = LoadPicture(App.Path & "\Botones\P" & Index & ".gif")
PintarBoton 0, 0
End Sub
Sub PintarBoton(Index As Integer, Posicion As Integer)
With PictureEnviar
.Picture = Boton(Index).Picture
.ForeColor = vbBlack
.CurrentX = 150 + Posicion
.CurrentY = 270 + Posicion
PictureEnviar.Print "Enviar"
.ForeColor = vbWhite
.CurrentX = 130 + Posicion
.CurrentY = 250 + Posicion
PictureEnviar.Print "Enviar"
End With
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
If Mascara Then
ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
Else
Call PopupMenu(Form2.mnuMascaras)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Boton(2).Visible = False
UnaVes = True
Eventos (X)
End Sub



Private Sub Form_Resize()

On Error Resume Next

Dim Posicion As POINTAPI
'GetCursorPos Posicion
If Me.Width < 10800 Then

mouse_event MOUSEEVENTF_LEFTUP, 1, 1, 0, 0
Me.Width = 10800
End If
If Me.Height < 8500 Then

mouse_event MOUSEEVENTF_LEFTUP, Posicion.X, Posicion.y, 0, 0
Me.Height = 8500

End If



moveSplit (1)
TileBackground

For i = 0 To 2
Call Redondear(Borde(i))
Next
Call Redondear(Picture2)
For i = 0 To 1
    Marco(i).Left = Me.ScaleWidth - 2850
    Foto(i).Left = Me.ScaleWidth - 2700
Next
Comando(0).Left = Me.ScaleWidth - 1200
Comando(1).Left = Me.ScaleWidth - 800
Comando(2).Left = Me.ScaleWidth - 400
Comando(3).Left = RichTextBox1.Width - 300

Marco(1).Top = Me.Height - 4100
Foto(1).Top = Me.Height - 3650
If Me.WindowState = 2 Then
Comando(4).Visible = False
Else
Comando(4).Visible = True
End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
Call SaveSetting(App.EXEName, "Color", "ColorFuente", RichTextBox2.SelColor)
Call SaveSetting(App.EXEName, "Tamaño", "TamañoFuente", RichTextBox2.SelFontSize)
Call SaveSetting(App.EXEName, "Fuente", "NombreFuente", RichTextBox2.SelFontName)
Call CloseDevice
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
RichTextBox1.Text = ""
Unload Form2
Unload Form3
Unload Form4
End Sub

Private Sub Image1_Click(Index As Integer)
On Error GoTo ReportedeError
Select Case Index
Case 0
Frame2.Visible = False
ListView1.ListItems.Clear

Case 1

Clipboard.Clear
Clipboard.SetData ImageList4.ListImages(ListView1.SelectedItem.Index).Picture
SendMessage RichTextBox2.hWnd, WM_PASTE, 0, 0
Case 2

Kill ImageList4.ListImages(ListView1.SelectedItem.Index).Key
CargarEmoticons

Case 4
Dim Extenciones As String, Abreviacion As String
Extenciones = "Todos los archivos de imágenes" + Chr$(0) + "*.gif" & ";" & "*.jpg" & ";" & "*.jpe" & ";" & "*.bmp" & ";" & "*.ico" + Chr$(0) + "Imágenes GIF (*.gif)" + Chr$(0) + "*.gif" + Chr$(0) + "Imágenes JPG (*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "Imágenes de mapas de bits (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "Iconos (*.ico)" + Chr$(0) + "*.ico" + Chr$(0) + "Todos los archivos (*.*)" + Chr$(0) + "*.*"

    strRutaImagen = ShowOpen(Extenciones)
    If strRutaImagen = "" Then Exit Sub
    strRutaImagen = Left(strRutaImagen, Len(strRutaImagen) - 1) 'no se porque mierda
    Abreviacion = Remplazar(InputBox("Elija la abreviación a utilizar"), True)
    If Abreviacion = "" Then Exit Sub
    FileCopy strRutaImagen, App.Path & "\emoticons\" & Abreviacion & Right(strRutaImagen, 4)
    CargarEmoticons

End Select
Exit Sub
ReportedeError:

MsgBox "Error " & err.Number & " " & err.Description
End Sub

Private Sub Label1_Change(Index As Integer)
Label1(1).Caption = Label1(0).Caption
End Sub



Private Sub Label3_Change(Index As Integer)
Label3(1).Caption = Label3(0).Caption

End Sub




Private Sub Marco_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then Exit Sub
CargarImagen Foto(1)
If Fotito.Width = 0 Then Exit Sub 'compruevo si se cargo alguna imagen
CentrarPicture Foto(1)
SavePicture Foto(1).Image, App.Path & "\FotoTemporal.bmp"
Enviar ("#Foto#")
End Sub


Sub CargarImagen(Cuadro As PictureBox)
On Error GoTo Descripcion
Dim sFile As String, Extenciones As String

Extenciones = "Todos los archivos de imágenes" + Chr$(0) + "*.gif" & ";" & "*.jpg" & ";" & "*.jpe" & ";" & "*.bmp" & ";" & "*.ico" + Chr$(0) + "Imágenes GIF (*.gif)" + Chr$(0) + "*.gif" + Chr$(0) + "Imágenes JPG (*.jpg)" + Chr$(0) + "*.jpg" + Chr$(0) + "Imágenes de mapas de bits (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0) + "Iconos (*.ico)" + Chr$(0) + "*.ico" + Chr$(0) + "Todos los archivos (*.*)" + Chr$(0) + "*.*"
sFile = ShowOpen(Extenciones)

If sFile <> "" Then
    Cuadro.Cls
    Set Fotito = LoadPicture(sFile)
    ancho = Fotito.Width
    alto = Fotito.Height
If ancho < Cuadro.Width And alto < Cuadro.Height Then porcentaje = 100: Exit Sub
        If ancho > Cuadro.Width Or alto > Cuadro.Height Then
            If ancho > alto Then
                porcentaje = (Cuadro.Width * 100) / ancho
            Else
                porcentaje = (Cuadro.Height * 100) / alto
            End If
        'CentrarPicture
        Exit Sub
        End If

    If ancho <= Cuadro.Width Or alto <= Cuadro.Height Then
        If ancho > alto Then
            porcentaje = (Cuadro.Width * 100) / ancho
        Else
            porcentaje = (Cuadro.Width * 100) / alto
        End If
    'CentrarPicture
    Exit Sub
    End If
End If
Exit Sub
Descripcion:
MsgBox "Error " & err.Number & " " & err.Description
End Sub

Public Sub CentrarPicture(Cuadro As PictureBox)
On Error Resume Next
Dim centro1 As Single, centro2 As Single
ancho = (ancho * porcentaje) / 100
alto = (alto * porcentaje) / 100
centro1 = (Cuadro.Width - ancho) / 2
centro2 = (Cuadro.Height - alto) / 2
Cuadro.PaintPicture Fotito, centro1, centro2, ancho, alto
Set Fotito = Nothing
End Sub


Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo ErrHandler
Dim ViejoNombre As String, NuevoNombre As String
ViejoNombre = ImageList4.ListImages(ListView1.SelectedItem.Index).Key
NuevoNombre = Replace(ViejoNombre, ImageList4.ListImages(ListView1.SelectedItem.Index).Tag, Remplazar(NewString, True))
Name ViejoNombre As NuevoNombre
CargarEmoticons
Exit Sub
ErrHandler:
MsgBox Error
End Sub

Private Sub ListView1_DblClick()
Image1_Click (1)
End Sub
Private Sub PictureEnviar_Click()
Send
End Sub

Private Sub PictureEnviar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
PintarBoton 1, 20
End Sub

Private Sub PictureEnviar_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
PintarBoton 0, 0
End Sub



Private Sub picVolumen_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 If (Button = vbLeftButton) Then
picVolumen.Left = picVolumen.Left + X
If picVolumen.Left <= Line1.X1 Then ' Or picVolumen.Left >= Line1.X2
mouse_event MOUSEEVENTF_LEFTUP, 1, 1, 0, 0
picVolumen.Left = Line1.X1 + 1
End If

If picVolumen.Left >= Line1.X2 Then
mouse_event MOUSEEVENTF_LEFTUP, 1, 1, 0, 0
picVolumen.Left = Line1.X2 - 100
End If

'Dim a As Long
 '   Dim tmp
    'vol = (picVolumen.Left - Line1.X1) * 5 - 5000
    'tmp = Right((Hex$(vol + 65536)), 4)
    'vol = CLng("&H" & tmp & tmp)
    'a = waveOutSetVolume(0, vol)
    '
End If
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
DoEvents
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0

Foto(1).PaintPicture Clipboard.GetData, 0, 0, Foto(1).Width, Foto(1).Height
SavePicture Foto(1).Image, App.Path & "\temporal1.bmp"
Dim Tamaño As Long
Open App.Path & "\temporal1.bmp" For Binary Access Read As #1
Tamaño = LOF(1)
ReDim Imagen(Tamaño - 1)
Get #1, , Imagen
Close #1
Winsock2.SendData Imagen
Winsock2.SendData "Fin"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandler
'On Error Resume Next
  Select Case Button.Key
    
    Case Is = "bold"
        If RichTextBox2.SelBold = True Then
           RichTextBox2.SelBold = False
        Else
           RichTextBox2.SelBold = True
        End If
        UpdateTextInfo
    Case Is = "italic"
        If RichTextBox2.SelItalic = True Then
           RichTextBox2.SelItalic = False
        Else
           RichTextBox2.SelItalic = True
        End If
        UpdateTextInfo
    Case Is = "underline"
        If RichTextBox2.SelUnderline = True Then
           RichTextBox2.SelUnderline = False
        Else
           RichTextBox2.SelUnderline = True
        End If
        UpdateTextInfo
    
    Case Is = "cut"
        SendKeys "+{DEL}"
        UpdateTextInfo
        
    Case Is = "copy"
        SendKeys "^{INSERT}"
        UpdateTextInfo
        
    Case Is = "paste"
        SendKeys "+{INSERT}"
        UpdateTextInfo
        
    Case Is = "Undo"
        SendKeys "^z"
        UpdateTextInfo
    Case Is = "imagen"
    InsertarImagen
      Case Is = "Objeto"
      InsertarObjeto
      Case Is = "Caritas"
      Frame2.Visible = True
      Frame2.ZOrder (0)
      CargarEmoticons
    Case "Sumbido"
    Enviar ("#Sumbido#")
  End Select
  Exit Sub
ErrHandler:
MsgBox Error
End Sub

Sub CargarEmoticons()
On Error Resume Next
Dim icon As String, Ruta As String
icon = Dir(App.Path & "\emoticons\*.*")
ListView1.ListItems.Clear
ListView1.Icons = Nothing
ImageList4.ListImages.Clear
While icon <> ""
Ruta = App.Path & "\emoticons\" & icon
ImageList4.ListImages.Add , Ruta, LoadPicture(Ruta)
ImageList4.ListImages(ImageList4.ListImages.Count).Tag = Remplazar(Left(icon, Len(icon) - 4), False)
icon = Dir
Wend
      

ListView1.Icons = ImageList4
ListView1.View = lvwIcon
For i = 1 To ImageList4.ListImages.Count
ListView1.ListItems.Add , , ImageList4.ListImages(i).Tag, i
Next
ListView1.PictureAlignment = lvwTopRight
End Sub

Private Sub Winsock1_Close()

RichTextBox2.Enabled = False: Frame1.BackColor = &H8000000F
Bootons_Click (0)
TextInfo NombreUsuario & " se a desconectado"
End Sub
Private Sub Winsock1_Connect()
Me.Show
MiNick = Form2.Text1(1)
Enviar (MiNick)
Unload Form2
RichTextBox2.Enabled = True: Frame1.BackColor = vbWhite

End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Me.Show
Winsock1.Close
Winsock1.Accept requestID
MiNick = Form2.Text1(1)
Enviar (MiNick)
Unload Form2
RichTextBox2.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo salir
Dim Dato As String
Winsock1.GetData Dato
If Not User Then NombreUsuario = Dato: User = True: Label1(0).Caption = NombreUsuario: Exit Sub
If NombreAdj Then NombreAdj = False: NombreAdjunto = Dato: TamañoAdj = True: Enviar ("#TamañoAdjunto#"): Exit Sub
If TamañoAdj Then TamañoAdj = False: TamañoAdjunto = Dato: InformarEnvio: Exit Sub
'-------

'--------
RichTextBox1.SelStart = Len(RichTextBox1.Text)

Select Case Dato

Case "#WebCam#"
SlideForm Form3, 0
TextInfo NombreUsuario & " quiere mostrarse por camara"
Exit Sub

Case "#AceptoWebCam#"
If Not Winsock2.State = 7 Then
Winsock2.CONNECT Winsock1.RemoteHostIP, 1000
End If
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hWnd, 0)
DoEvents: SendMessage mCapHwnd, CONNECT, 0, 0
Timer1.Enabled = True
TextInfo NombreUsuario & "a aceptado tu imbitacion para ver tu WebCam"
Exit Sub

Case "#FinWebCam#"
TextInfo NombreUsuario & " a desconectado la WebCam"
Enviar ("#AceptoFoto#") ' cuando el otro desconecta su webcam yo pido ver su foto
Exit Sub

Case "#Audio#"
SlideForm Form4, 0
TextInfo NombreUsuario & " lo imbita a escuchar su vos"
Exit Sub


Case "#AceptoAudio#"
If Not Winsock3.State = 7 Then
Winsock3.CONNECT Winsock1.RemoteHostIP, 1001
End If
Picture2.Visible = True
RECORD_Start
Call OpenDevice
Timer2.Enabled = True
TextInfo NombreUsuario & "acepto tu imbitacion para oir tu vos"

Case "#FinAudio#"
TextInfo NombreUsuario & " a desconectado su microfono"
Close #2
picVolumen.Visible = False
Line1.Visible = False
Exit Sub

Case "#Foto#"
On Error Resume Next
If Not Winsock2.State = 7 Then
Winsock2.LocalPort = 1000
Winsock2.Listen
End If
Open Ruta For Binary As #1
Enviar ("#AceptoFoto#")

Case "#AceptoFoto#"
If Not Winsock2.State = 7 Then
Winsock2.CONNECT Winsock1.RemoteHostIP, 1000
Else
EnviarFoto
End If

Case "#Adjunto#"
Enviar ("#NombreAdjunto#")
NombreAdj = True

Case "#NombreAdjunto#"
Enviar (NombreArchivo)

Case "#TamañoAdjunto#"
Enviar (TamañoArchivo)
TamañoArchivo = TamañoArchivo / 10
Case "#AceptoAdjunto#"
If Not Winsock4.State = 7 Then
Winsock4.CONNECT Winsock1.RemoteHostIP, 998
Else
EnviarArchivo
End If
Case "#CacelarAdjunto#"
RichTextBox1 = Replace(RichTextBox1, "Webdings", "Lucida Console")
RichTextBox1 = Replace(RichTextBox1, nn2 & "\cell\cf0\f", "    Cancelado" & "\cell")
RichTextBox1 = Replace(RichTextBox1, "Cancelar_", "")
vv2 = "": nn2 = "": Progreso2 = 0
TextInfo NombreUsuario & " a cancelado el envio de " & NombreArchivo
Adjunto = False
Winsock4.Close

Case "#Sumbido#"
Sumbido

Case "#Escriviendo#"
Label3(0).Caption = NombreUsuario & " te esta escriviendo un mensage."
Exit Sub

Case Else
If Me.WindowState = vbMinimized Then
sndplaysound "C:\WINDOWS\Media\chimes.wav", SND_NODEFAULT + SND_ASYNC  'App.Path &
FlashWindow Me.hWnd, Invert
End If

rtf = rtf & Dato
If Right(Dato, 4) = "ter3" Or Dato = "ter3" Then
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelRTF = rtf
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.Refresh
RichTextBox2.SetFocus
Label3(0).Caption = "Ultimo mensage recivido " & Now
rtf = ""
Exit Sub
End If
End Select
Exit Sub
salir:
If err.Number = 32755 Then Exit Sub ' cancelar del commandialog
MsgBox "Error " & err.Number & " " & err.Description
End Sub

Sub ComboFuentes()
With Picture1
.AutoRedraw = True
.BackColor = vbWhite
'.Visible = False
'---cambiar estos 3 valores para modificar el tamaño
.FontSize = 17 'Ej: 24
.Height = 330 'Ej: 500
.Width = 2895 'Ej: 3800
'---------
ImageCombo1.Height = .Height
ImageCombo1.Width = .Width
End With

For i = 0 To Screen.FontCount - 1
DoEvents
With Picture1
.Picture = Nothing
.FontName = Screen.Fonts(i)
.CurrentX = 0
.CurrentY = 0
Picture1.Print Screen.Fonts(i)
End With
ImageList2.ListImages.Add , Picture:=Picture1.Image
Set ImageCombo1.ImageList = ImageList2
ImageCombo1.ComboItems.Add , , Screen.Fonts(i), i + 1, i + 1
Next i
End Sub

Sub ComboColores()

Picture1.Height = 330
Picture1.Width = 1220
    For i = 0 To 15
    DoEvents
    Picture1.BackColor = QBColor(i)
    
    ImageList3.ListImages.Add , Picture:=Picture1.Image
    
    Set ImageCombo3.ImageList = ImageList3
ImageCombo3.ComboItems.Add , , QBColor(i), i + 1, i + 1

Next
End Sub
Sub ComboTamaño()
Dim intI As Integer
For intI = 8 To 72 Step 2
     Combo1.AddItem Str(intI)
Next intI

End Sub
Private Sub RichTextBox2_GotFocus()
UpdateTextInfo
End Sub
Sub UpdateTextInfo()
On Error GoTo eds
    
    Static fLast As String 'font
    Static sLast As String 'font size
    

        If RichTextBox2.SelBold Then
            Toolbar1.Buttons("bold").Value = tbrPressed
        Else
            Toolbar1.Buttons("bold").Value = tbrUnpressed
        End If
        
        
        If RichTextBox2.SelItalic Then
            Toolbar1.Buttons("italic").Value = tbrPressed
        Else
            Toolbar1.Buttons("italic").Value = tbrUnpressed
        End If
        
        
        If RichTextBox2.SelUnderline Then
            Toolbar1.Buttons("underline").Value = tbrPressed
        Else
            Toolbar1.Buttons("underline").Value = tbrUnpressed
        End If
        
        
        
        Dim blnFound As Boolean
        Dim intI As Integer ' counter
        If fLast <> RichTextBox2.SelFontName Then
           
            For intI = 1 To ImageCombo1.ComboItems.Count
            
                If ImageCombo1.ComboItems(intI) = RichTextBox2.SelFontName Then
                    ImageCombo1.SelectedItem = ImageCombo1.ComboItems(intI)
                    intI = ImageCombo1.ComboItems.Count
                    blnFound = True
                End If
            Next intI
            If blnFound = True Then
                fLast = RichTextBox2.SelFontName
            Else
               
                fLast = ""
                ImageCombo1.SelectedItem = ImageCombo1.ComboItems(-1)
            End If
        End If
        
    Combo1.Text = RichTextBox2.SelFontSize
    Dim Index As Integer
    For Index = 1 To 15
        If Val(ImageCombo3.ComboItems(Index).Text) = RichTextBox2.SelColor Then
    
            ImageCombo3.SelectedItem = ImageCombo3.ComboItems(Index)
            
            Exit For
        End If
    Next

eds:

End Sub

Private Sub ImageCombo3_Change()
RichTextBox2.SelColor = Val(ImageCombo3.Text)
End Sub
Private Sub ImageCombo3_Click()
RichTextBox2.SelColor = Val(ImageCombo3.Text)
End Sub
Private Sub ImageCombo1_Change()
RichTextBox2.SelFontName = ImageCombo1.Text
End Sub
Private Sub ImageCombo1_Click()
RichTextBox2.SelFontName = ImageCombo1.Text
End Sub
Private Sub Combo1_Change()
RichTextBox2.SelFontSize = Combo1.Text
End Sub
Private Sub Combo1_Click()
RichTextBox2.SelFontSize = Combo1.Text
End Sub
Private Sub RichTextBox2_SelChange()
UpdateTextInfo
End Sub
Sub RecuperarInfo()
With RichTextBox2
.Font = ImageCombo1.Text
.Font.Size = Combo1.Text
.SelColor = Val(ImageCombo3.Text)
.Font.Bold = Toolbar1.Buttons(1).Value
.Font.Italic = Toolbar1.Buttons(2).Value
.Font.Underline = Toolbar1.Buttons(3).Value
End With
End Sub

Private Sub RichTextBox2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 8
Send
Exit Sub
End If
If Not Escriviendo Then
Enviar ("#Escriviendo#")
Escriviendo = True
End If
End Sub
Private Sub picsplit_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim fMoveX As Single
    Select Case KeyCode
    Case vbKeyLeft
        fMoveX = -60
        If (Shift And vbShiftMask) = 1 Then fMoveX = -360
        If (Shift And vbCtrlMask) = 2 Then fMoveX = (90 - Me.picSplit.Top)
        Call moveSplit(fMoveX)
    Case vbKeyRight
        fMoveX = 60
        If (Shift And vbShiftMask) = 1 Then fMoveX = 360
        If (Shift And vbCtrlMask) = 2 Then fMoveX = (Me.Height - 120 - Me.picSplit.Top)
        Call moveSplit(fMoveX)
    Case vbKeyReturn, vbKeyTab
        If Me.RichTextBox1.Visible Then
            Me.RichTextBox1.SetFocus
        Else
            Me.RichTextBox2.SetFocus
        End If
    End Select
End Sub


Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If (Button = vbLeftButton) Then
        Call moveSplit(y)
        
    End If
End Sub

Private Sub moveSplit(ByVal vfMoveX As Single)
On Error Resume Next
    Dim fNewRightAreaPos As Single
    Dim fNewRightAreaHeight As Single
    Dim fNewSplitTopftPos As Single
   
    fNewRightAreaPos = Me.RichTextBox2.Top + vfMoveX
    fNewRightAreaHeight = Me.ScaleHeight - fNewRightAreaPos
    fNewSplitTopftPos = Me.picSplit.Top + vfMoveX
    
    If fNewRightAreaPos > 3000 And fNewRightAreaHeight > 1500 And fNewSplitTopftPos > 30 And fNewSplitTopftPos < (Me.ScaleHeight - 45) Then
        If Marco(0).Visible Then
        RichTextBox1.Move RichTextBox1.Left, RichTextBox1.Top, Me.ScaleWidth - 3150, picSplit.Top - picSplit.Height - 1250
        RichTextBox2.Move RichTextBox2.Left, fNewRightAreaPos, Me.ScaleWidth - 4500, fNewRightAreaHeight - 600
        Else
        RichTextBox1.Move RichTextBox1.Left, RichTextBox1.Top, Me.ScaleWidth - 400, picSplit.Top - picSplit.Height - 1250
        RichTextBox2.Move RichTextBox2.Left, fNewRightAreaPos, Me.ScaleWidth - 1750, fNewRightAreaHeight - 600
        End If
       RichTextBox2.Top = fNewRightAreaPos
       Toolbar1.Top = RichTextBox2.Top - Toolbar1.Height:  Toolbar1.Width = RichTextBox1.Width - 20
       Borde(1).Move Borde(1).Left, RichTextBox2.Top - 630, RichTextBox1.Width
       Borde(2).Move Borde(2).Left, Me.ScaleHeight - 650, RichTextBox1.Width
       Borde(0).Width = RichTextBox1.Width
       Frame1.Height = RichTextBox2.Height: Frame1.Left = RichTextBox1.Width - 1350: Frame1.Top = fNewRightAreaPos
       PictureEnviar.Top = Frame1.Height / 2 - PictureEnviar.Height / 2
    picSplit.Move Me.picSplit.Left, fNewSplitTopftPos, RichTextBox1.Width, picSplit.Height
    
    End If
End Sub
Private Sub RichTextBox2_Change()
Dim found As Integer
Dim emoticon As Integer

Dim Palabra As String
Dim LenPalabra As Integer
If Check1 Then
For i = 1 To ImageList4.ListImages.Count
Palabra = ImageList4.ListImages(i).Tag
LenPalabra = Len(Palabra)
If RichTextBox2.Find(Palabra, RichTextBox2.SelStart - LenPalabra, Len(RichTextBox2.Text)) > -1 Then
    found = RichTextBox2.Find(Palabra, RichTextBox2.SelStart - LenPalabra, Len(RichTextBox2.Text))
    emoticon = i
    SET_PICTURE found, emoticon, LenPalabra
End If
Next
End If

End Sub

Public Function SET_PICTURE(pos As Integer, emoticon As Integer, LenPalabra As Integer)
On Error Resume Next
Clipboard.Clear
Clipboard.SetData ImageList4.ListImages(emoticon).Picture
RichTextBox2.SelStart = pos
RichTextBox2.SelLength = LenPalabra
RichTextBox2.SelText = ""
SendMessage RichTextBox2.hWnd, WM_PASTE, 0, 0
End Function
Sub TileBackground()
Dim bgdImage    As Picture
Dim X           As Integer
Dim y           As Integer

'Used to tile bakground picture

Set bgdImage = Me.Picture
y = 0
While y < Me.Height
    X = 0
    While X < Me.Width
        PaintPicture bgdImage, X, y
        picSplit.PaintPicture bgdImage, X, y
        X = X + bgdImage.Width \ 2
    Wend
        y = y + bgdImage.Height \ 2
Wend





    X = 0
    While X < Me.Width
        PaintPicture Toolsbar, X, 0
        X = X + Toolsbar.Width \ 2
    Wend
    
 'Set bgdImage = LoadPicture(App.Path & "\Ventana.gif")
      
'PaintPicture LoadPicture(App.Path & "\Ventana.gif"), Me.ScaleWidth - 5000, 0


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AgregarIcono
MostrarGlobo (Description)
End Sub

Private Sub Winsock2_Connect()
EnviarFoto
End Sub
Sub EnviarFoto()
On Error Resume Next
Dim Tamaño As Long
If WebCam Then
Open App.Path & "\temporal1.bmp" For Binary Access Read As #5
Else
Open App.Path & "\FotoTemporal.bmp" For Binary Access Read As #5
End If
Tamaño = LOF(5)
ReDim Imagen(Tamaño - 1)
Get #5, , Imagen
Close #5
Winsock2.SendData Imagen
Winsock2.SendData "Fin"
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept requestID
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Imagen As String
Winsock2.GetData Imagen, vbNullString
Put #1, , Imagen
If Len(Replace(Imagen, "Fin", "")) < Len(Imagen) Then
Put #1, , Replace(Imagen, "Fin", "")
Close #1
Foto(0).Picture = LoadPicture(Ruta)
Open Ruta For Binary As #1
End If
End Sub
Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
vol = (picVolumen.Left - Line1.X1) * 2 - 2000
'On Error Resume Next
Dim Audio As String
Winsock3.GetData Audio, vbNullString
If Len(Replace(Audio, "Stop", "")) < Len(Audio) Then
Put #2, , Replace(Audio, "Stop", "")
Close #2
'---------------
Set DSToneBuffer = DS.CreateSoundBufferFromFile(RutaAudio, desc)
DSToneBuffer.SetVolume Val(vol)
DSToneBuffer.Play DSBPLAY_DEFAULT

'sndplaysound RutaAudio, SND_NODEFAULT + SND_ASYNC
Open RutaAudio For Binary As #2
Else
Put #2, , Audio
End If
End Sub
Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Winsock3.Close
Winsock3.Accept requestID
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Static Contador
Contador = Contador + 1


If Contador = 50 Then
RECORD_Save
RECORD_Start
Open "c:\TempWave.wav" For Binary Access Read As #4
Dim Tamaño As Long
Tamaño = LOF(4)
ReDim AudioB(Tamaño - 1)
Get #4, , AudioB
Close #4
Winsock3.SendData AudioB
Winsock3.SendData "Stop"
Contador = 0
End If
GraficarAudio
End Sub

Private Sub Winsock4_Connect()
EnviarArchivo
End Sub
Private Sub Winsock4_ConnectionRequest(ByVal requestID As Long)
Winsock4.Close
Winsock4.Accept requestID
End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Accion
Progreso = Progreso + bytesTotal
nn = String(Progreso / TamañoAdjunto, "g")
If nn <> vv Then Graficar
Dim Archivo As String
Winsock4.GetData Archivo, vbNullString
If Len(Replace(Archivo, "#FinArchivo#", "")) < Len(Archivo) Then
Put #6, , Replace(Archivo, "#FinArchivo#", "")
Close #6
RichTextBox1 = Replace(RichTextBox1, "Webdings", "Lucida Console")
RichTextBox1 = Replace(RichTextBox1, vv & "\cell\cf0\f", "Descarga Completa" & "\cell")
'RutaGuardar = GetShortPath(RutaGuardar)
RutaGuardar = Replace(GetShortPath(RutaGuardar), "\", "\\")
RichTextBox1 = Replace(RichTextBox1, "_Cancelar_", "")
nn = "": vv = "": Progreso = 0
TextInfo "Finalizo la transferencia de " & NombreAdjunto
TextOkCancel "", ""
TextOkCancel RutaGuardar, ""
RichTextBox1.SelStart = Len(RichTextBox1.Text)
Else
Put #6, , Archivo
End If
Exit Sub
Accion:
Close #6
Winsock4.Close
MsgBox Error
End Sub
Sub Graficar()
'DoEvents
RichTextBox1 = Replace(RichTextBox1, vv & "\cell\cf0\f", nn & "\cell\cf0\f")
vv = nn
RichTextBox1.SelStart = Len(RichTextBox1.Text)
End Sub

Sub EnviarArchivo()
Open RutaArchivo For Binary Access Read As #7
Dim Tamaño As Long
Tamaño = LOF(7)
ReDim Archivo(Tamaño - 1)
Get #7, , Archivo
Close #7
Winsock4.SendData Archivo
Winsock4.SendData "#FinArchivo#"
End Sub
Sub Sumbido()
Dim X As Single, y As Single
Dim lngWindowPosition As Long
If Me.WindowState = 1 Then Me.WindowState = 0
lngWindowPosition = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
X = Me.Top
y = Me.Left
sndplaysound "C:\Archivos de programa\MSN Messenger\nudge.wav", SND_NODEFAULT + SND_ASYNC  'App.Path &
For i = 1 To 100
DoEvents
Me.Top = X + Rnd(1000) * 150
Me.Left = y + Rnd(1000) * 150
Next
lngWindowPosition = SetWindowPos(Me.hWnd, HWND_NOTTOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
TextInfo NombreUsuario & " te envio un sumbido"
End Sub
Sub Enviar(Cadena As String)
On Error GoTo Descripcion

Winsock1.SendData Cadena
Exit Sub
Descripcion:
MsgBox "Error " & err.Number & " " & err.Description
End Sub

Sub MostrarFoto()
On Error Resume Next
Enviar ("#FinWebCam#")
Foto(1).Cls
Foto(1).PaintPicture LoadPicture(App.Path & "\FotoTemporal.bmp"), 0, 0, Foto(1).Width, Foto(1).Height
End Sub

Sub VerSuWebCam()
On Error Resume Next
If Not Winsock2.State = 7 Then
Winsock2.LocalPort = 1000
Winsock2.Listen
End If
Open Ruta For Binary As #1
Enviar ("#AceptoWebCam#")
End Sub

Sub EscucharSuAudio()
If Not Winsock3.State = 7 Then
Winsock3.LocalPort = 1001
Winsock3.Listen
End If
Open RutaAudio For Binary As #2
Enviar ("#AceptoAudio#")
picVolumen.Visible = True
Line1.Visible = True
End Sub
Sub AceptarAdjunto()
Dim sFile As String
TamañoAdjunto = TamañoAdjunto / 10
sFile = Guardar
If sFile <> "" Then
If Not Right(sFile, 1) = "\" Then sFile = sFile & "\"
RutaGuardar = sFile & NombreAdjunto
Open sFile & NombreAdjunto For Binary As #6
If Not Winsock4.State = 7 Then
Winsock4.LocalPort = 998
Winsock4.Listen
End If
RichTextBox1 = Replace(RichTextBox1, "_Aceptar_\ulnone           ", "")

'RichTextBox1.SelStart = Len(RichTextBox1.Text)
'-------

Enviar ("#AceptoAdjunto#")

End If
'End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Existe(Text) Then
ShellExecute 0, vbNullString, Text, vbNullString, vbNullString, 1
End If
Select Case Text
Case "_Cancelar_"
Enviar "#CacelarAdjunto#"
RichTextBox1 = Replace(RichTextBox1, "Webdings", "Lucida Console")
RichTextBox1 = Replace(RichTextBox1, vv & "\cell\cf0\f", "    Cancelado" & "\cell")
RichTextBox1 = Replace(RichTextBox1, "_Cancelar_", "")
RichTextBox1 = Replace(RichTextBox1, "_Aceptar_", "")
RichTextBox1.SelStart = Len(RichTextBox1.Text)
Winsock4.Close
Close #6
nn = "": vv = "": Progreso = 0
Case "_Aceptar_"
AceptarAdjunto
End Select
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, y As Single)
    
    
    Text = GetWord(RichTextBox1, X, y)
    'If Label4.Caption <> Text Then Label4.Caption = Text
    If Existe(Text) Then
    RichTextBox1.MousePointer = 99
    Else
    RichTextBox1.MousePointer = 1
    End If
    
End Sub

Sub InformarEnvio()
TextInfo (NombreUsuario & " le quiere enviar un archivo.")
Open "Temp" & NombreAdjunto For Binary As #6
Close #6
PegarIcono "Temp" & NombreAdjunto
Kill "Temp" & NombreAdjunto
TextOkCancel NombreAdjunto, "(" & Trim(Format$(Format$((TamañoAdjunto \ 1024) + 1, "##,###,##0") & " KB)", "@@@@@@@@@@@@"))
RichTextBox1.SelRTF = Text1
TextOkCancel "_Aceptar_", "_Cancelar_"
'RichTextBox1.SelStart = Len(RichTextBox1.Text)
End Sub
Sub PegarIcono(RutaIcono As String)
Dim FI As SHFILEINFO
Dim hImage As Long
Picture1.Width = 520: Picture1.Height = 520
Picture1.Cls
hImage = SHGetFileInfo(RutaIcono, ByVal 0&, FI, Len(FI), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
ImageList_Draw hImage, FI.iIcon, Picture1.hDC, 0, 0, ILD_TRANSPARENT
Clipboard.Clear
Clipboard.SetData Picture1.Image
RichTextBox1.Locked = False
RichTextBox1.SelText = Space(5)
SendMessage RichTextBox1.hWnd, WM_PASTE, 0, 50
RichTextBox1.Locked = True
RichTextBox1.SelText = vbCrLf
End Sub
Private Sub Winsock4_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'On Error Resume Next
'MsgBox bytesSent
Progreso2 = Progreso2 + bytesSent
nn2 = String(Progreso2 / TamañoArchivo, "<")

If nn2 <> vv2 Then
RichTextBox1 = Replace(RichTextBox1, vv2 & "\cell\cf0\f", nn2 & "\cell\cf0\f")
vv2 = nn2
RichTextBox1.SelStart = Len(RichTextBox1.Text)
End If
If Len(nn2) >= Len("<<<<<<<<<<") Then
RichTextBox1 = Replace(RichTextBox1, "Webdings", "Lucida Console")
RichTextBox1 = Replace(RichTextBox1, nn2 & "\cell\cf0\f", " Envio Completo" & "\cell")
RichTextBox1 = Replace(RichTextBox1, "Cancelar_", "")
vv2 = "": nn2 = "": Progreso2 = 0
TextInfo "Se completo el envio de " & NombreArchivo
Adjunto = False
End If
Exit Sub
'salir:
End Sub
