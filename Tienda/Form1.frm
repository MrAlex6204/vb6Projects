VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Ventas"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   14925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3135
      Left            =   480
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   9975
      Begin VB.CommandButton Command7 
         Caption         =   "X"
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
         Left            =   9240
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         ItemData        =   "Form1.frx":0000
         Left            =   3960
         List            =   "Form1.frx":0013
         TabIndex        =   36
         Top             =   840
         Width           =   4335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
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
         Left            =   1440
         TabIndex        =   35
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text13 
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
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Aceptar"
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
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   435
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   2100
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   435
         Left            =   4200
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   3960
         TabIndex        =   33
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label LabelTot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "COBRANZA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   5400
      Width           =   9975
      Begin VB.CommandButton Command8 
         Caption         =   "&Consultar"
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
         Left            =   4080
         TabIndex        =   39
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Total"
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
         Left            =   1440
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblConsulta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   435
         Left            =   3240
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num Art"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   870
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Consulta y Edicion de Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   10560
      TabIndex        =   20
      Top             =   960
      Width           =   4695
      Begin VB.CommandButton Command11 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         DataField       =   "Descrip"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text10 
         DataField       =   "NumArt"
         DataSource      =   "Adodc1"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         DataField       =   "Precio"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   3135
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   240
         Top             =   3120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"Form1.frx":005C
         OLEDBString     =   $"Form1.frx":00F1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ListArt"
         Caption         =   "Articulos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num Art"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Captura de Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   10560
      TabIndex        =   8
      Top             =   1680
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text5 
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
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num Art"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1245
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   4380
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":0186
      Left            =   480
      List            =   "Form1.frx":018D
      MousePointer    =   15  'Size All
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   840
      Width           =   9975
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   435
      Left            =   4560
      TabIndex        =   19
      Top             =   240
      Width           =   2100
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
      Begin VB.Menu menuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuArticulos 
      Caption         =   "Articulos"
      Begin VB.Menu mnuList 
         Caption         =   "Lista De Articulos"
      End
   End
   Begin VB.Menu mnuCierracaja 
      Caption         =   "Cierre de Caja"
      Begin VB.Menu mnuCorte 
         Caption         =   "Corte de Caja"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDD As Database 'Objeto para manejar la base de datos
Dim TBL As Recordset  'Objeto para manejar la Tabla
Dim SQL As String
   
   
   Dim sItemData As String
   Dim strData As String 'Variable Que Sirvira Para Localizar Un Dato
   Dim strOutData As String
   Dim strConnect As String
   Public RS As New ADODB.Recordset
   Public RVentArt As New ADODB.Recordset
   Dim strPath As String
   Dim articulos As Integer
   Dim total As Double
   Dim Dia As Date
   Dim tipCambio As String
   
 '  Public RS As New Recordset

Public Sub BD()



End Sub

 Sub Command1_Click()
 
 On Error Resume Next
 If Text1 = Empty Then
 Beep
 Exit Sub
 End If
 
AbrirBase

'Esta Variable Contiene El Campo y el Valor a Buscar Dentro de la Tabla _
Nota Lo que Va entre comillas es el Dato a Buscar
 strData = "NumArt = '" & Text1.Text & "'"
'El RS.Find strData sirve Para EContrar Lo que hay en la variable strData

    RS.Find strData
    If RS.EOF = True Then
    
    If MsgBox("NO SE ENCONTRO ARTICULO !!! ¿DESEA AGREGARLO ?", vbYesNo, "ALTA PRODUCTO") = 6 Then
     Call Frame2_Click
    End If
    
    
    CerrarBase
    Exit Sub
    End If
    
   
     Text2 = RS.Fields("Precio")
     Text3 = RS.Fields("Descrip")
     RS.MoveFirst
     
     List2.AddItem (Text1 + "   " + Text3 + "   " + Text2)
     articulos = articulos + 1
    
    
    
    


     Text1.SetFocus
     
   
       
       RVentArt.AddNew
       RVentArt.Fields("Cajero") = frmLogin.NomCajero
       RVentArt.Fields("NumArt") = Text1
       RVentArt.Fields("PrecArt") = Text2
       RVentArt.Update
       
       total = total + Text2
       
     CerrarBase
    
    
    
   
     
Text1 = Empty
Text1.SetFocus
End Sub

Sub Command10_Click()

If Text13 = Empty Then
Beep
Exit Sub
End If

lblTotal = Val(lblTotal) - Val(Text13)

If Val(lblTotal) <= 0 Then
Label8.Visible = True
Label8 = "Cambio:" + Str(Abs(Text12))

Command3.Visible = True
Command10.Enabled = False
Text13.Enabled = False
Exit Sub
End If


End Sub

Private Sub Command11_Click()

Frame4.Height = 495
Frame4.Width = 4695


End Sub

Private Sub Command2_Click()

If Text6 = Empty Or Text5 = Empty Or Text4 = Empty Then
MsgBox "Favor de Llenar los Todos los Datos", vbCritical, "VeraSoft"
Exit Sub
End If

AbrirBase
RS.AddNew
RS.Fields("Precio") = Text6
RS.Fields("Descrip") = Text4
RS.Fields("NumArt") = Text5
RS.Update

Text5 = Empty
Text6 = Empty
Text4 = emty

CerrarBase
End Sub

Private Sub Command3_Click()
If tipCambio = Empty Then
MsgBox "Seleccione Forma de Pago", vbCritical, "VeraSoft"
Exit Sub
End If


List2.Clear

RS.Open "select * from VentReg", strConnect, adOpenKeyset, adLockOptimistic

RS.AddNew
RS.Fields("cajero") = frmLogin.NomCajero
RS.Fields("Total") = total
RS.Fields("Fecha") = Dia
RS.Fields("TipCambio") = tipCambio
RS.Update

RS.Close

'Resetea Todos Los Valores de Los Controles
Label8.Visible = False
Command3.Visible = False
Label14 = "Abono:"
Frame5.Visible = False
Text13.Enabled = True
Text13 = Empty
total = 0
Command10.Enabled = True



End Sub

 Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()

Frame2.Height = 495
Frame2.Width = 4695
End Sub

Private Sub Command7_Click()
Frame5.Visible = False
End Sub

 Sub Command8_Click()
 lblConsulta.Visible = True
AbrirBase


 RS.Find "NumArt = '" & Text1.Text & "'"
 
    If RS.EOF = True Then
    
    If MsgBox("NO SE ENCONTRO ARTICULO !!! ¿DESEA AGREGARLO ?", vbYesNo, "ALTA PRODUCTO") = 6 Then
     Call Frame2_Click
    End If
    
    
    CerrarBase
    Exit Sub
    End If
    
    
    lblConsulta.Caption = RS.Fields("Descrip") + "----->" + Str(RS.Fields("Precio"))
   
     
     RS.MoveFirst
     
     
CerrarBase
End Sub

Sub Command9_Click()


End Sub



Private Sub DataGrid1_DblClick()
DataGrid1.Visible = False
Adodc1.ConnectionString = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Tienda.mdb;Mode=ReadWrite;Persist Security Info=False"
Adodc1.Refresh
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
Adodc1.ConnectionString = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Tienda.mdb;Mode=ReadWrite;Persist Security Info=False"
DataGrid1.Visible = False
Text1.SetFocus
End If
End Sub

Private Sub Command6_Click()
Frame5.Visible = True

If total = 0 Then
MsgBox "No hay Articulos Cobrados", vbCritical, "VeraSoft"
Exit Sub
End If

lblTotal = total

End Sub

Private Sub Form_Load()
total = 0
articulos = 0
Dia = Date
List2.Clear
tipCambio = Empty

Adodc1.ConnectionString = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Tienda.mdb;Mode=ReadWrite;Persist Security Info=False"

Frame1.Caption = Frame1.Caption + " LE ATIENDE EL CAJERO: " + frmLogin.NomCajero
    'Abre la base de datos
    Set BDD = OpenDatabase(App.Path & "\Tienda.mdb")
    
   
     
    'Obtiene el Path de la Base de Datos
    strPath = App.Path & "\Tienda.mdb"
    
    'Variable Que Contine El Tipo de Coenccion que se va realizar
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
    
     
     
    
End Sub
'Metodo Para Abrir La Base De Datos
Sub AbrirBase()
RS.Open "select * from ListArt", strConnect, adOpenKeyset, adLockOptimistic
RVentArt.Open "select * from VentArt", strConnect, adOpenKeyset, adLockOptimistic
End Sub

'Metodo Para Cerrar la Base de Datos
Sub CerrarBase()
RS.Close
RVentArt.Close
End Sub

Private Sub Form_Terminate()
Adodc1.Refresh
End Sub

Private Sub Frame2_Click()


Frame2.Height = 4095
Frame2.Width = 4335
End Sub

Private Sub Frame3_Click()


Frame3.Height = 4095
Frame3.Width = 4335
End Sub

Private Sub Frame4_Click()




Frame4.Height = 4595
Frame4.Width = 4335

End Sub

Private Sub Label7_Click()
Label7.Caption = frmLogin.NomCajero
End Sub

Private Sub List1_Click()
If List1.ListIndex <> -1 Then
   'para pbtener el dato de un ittem seleccionado

 tipCambio = List1.List(List1.ListIndex)
  

End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If List1.ListIndex <> -1 Then
   'para pbtener el dato de un ittem seleccionado
  Text1 = List1.List(List1.ListIndex)
  
  
  

End If
End If
End Sub

Private Sub List2_Click()

If List2.ListIndex <> -1 Then
   'para pbtener el dato de un ittem seleccionado

  lblDescrip.Caption = List2.List(List2.ListIndex)

  
End If


End Sub

Private Sub menuCerrar_Click()
Me.Hide
End Sub

Private Sub mnuCorte_Click()
Dialog1.Show
End Sub

Private Sub mnuList_Click()
Dialog.Show
End Sub

Private Sub Text1_Click()
lblConsulta.Visible = False
lblDescrip.Caption = "Descripcion"

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 112 Then
Call Command3_Click
DataGrid1.SetFocus
End If
Label12 = KeyCode
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

Call Command1_Click

End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command10_Click
Text13 = Empty
End If
End Sub
