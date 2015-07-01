VERSION 5.00
Object = "{40BD39E3-6F1E-11D1-B2DF-444553540000}#1.0#0"; "SHAPE.OCX"
Begin VB.Form Form1 
   Caption         =   "Custom Form Shape activeX Demo"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "Demo2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   StartUpPosition =   3  'Windows Default
   Begin FormShape.FormShape FormShape1 
      Left            =   3360
      Top             =   3120
      ShapeType       =   1
      MaskColor       =   0
      AutoScale       =   -1  'True
      ScaleX          =   1
      ScaleY          =   1
      ShapePicture    =   "Demo2.frx":030A
      ShapeString     =   "{(0,0)(100,0)[200,0,-180](400,0)[200,0,180]}{(150,0)[200,0,-180]}"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resize the Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   3450
   End
   Begin FormShape.FormShape FormShape3 
      Left            =   720
      Top             =   480
      ShapeType       =   0
      MaskColor       =   0
      AutoScale       =   -1  'True
      ScaleX          =   1
      ScaleY          =   1
      ShapeString     =   "{(0,0)(100,0)[200,0,-180](400,0)[200,0,180]}{(150,0)[200,0,-180]}"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    FormShape1.hWnd = Form1.hWnd
End Sub

Private Sub Form_Resize()
    FormShape1.Refresh
End Sub
