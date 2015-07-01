VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Deaclaraciones del Api

'Función que sirve para creae una región eliptica
Private Declare Function CreateEllipticRgn Lib "gdi32" ( _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

' SetWindowRgn signa la región anterior a la ventana
Private Declare Function SetWindowRgn Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hRgn As Long, _
        ByVal bRedraw As Long) As Long
        

'El primer parámetro es el objeto Formualario, luego las dimensiones
Private Sub Form_Redondo( _
            Objeto As Object, _
            x As Long, y As Long, _
            Ancho As Long, Alto As Long)
            
' Hnadle para la región devuelta por CreateEllipticRgn
Dim Region As Long
    
    
    Objeto.BackColor = vbRed
    
    'Crea la ergión elíptica
    Region = CreateEllipticRgn(x, y, _
             Objeto.ScaleX(Ancho, Objeto.ScaleMode, vbPixels), _
             Objeto.ScaleY(Alto, Objeto.ScaleMode, vbPixels))
    
    ' Establece la región a la ventana
    Call SetWindowRgn(Objeto.hwnd, Region, True)
    
    
    With Objeto
       .Cls
       .AutoRedraw = True
       .FontSize = 26
       .ForeColor = vbWhite
       .ScaleMode = 1
       .CurrentX = (.Width - .TextWidth("..Wenas ")) / 2
       .CurrentY = (.Height - .TextHeight("..Wenas ")) / 2
       
       Objeto.Print "..Wenas "
    End With
    

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
    Call Form_Redondo(Me, 5, 5, Me.ScaleWidth, Me.ScaleHeight)
End Sub

