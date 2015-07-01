Attribute VB_Name = "Calc_Picture"




Option Explicit

Public CalculationDone As Boolean
Public TransColor As Long
Public ByteCtr As Long
Public RgnData() As Byte

Private Const RGN_XOR = 3
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long


Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long



Private PicInfo As BITMAP

Private Type BITMAP
bmType As Long
bmWidth As Long
bmHeight As Long
bmWidthBytes As Long
bmPlanes As Integer
bmBitsPixel As Integer
bmBits As Long
End Type

'Calculate a Region to shape the form
Public Sub CalcPic(Pic As PictureBox)

Dim rgnMain As Long
Dim X As Long
Dim Y As Long
Dim rgnPixel As Long
Dim RGBColor As Long
Dim dcMain As Long
Dim bmpMain As Long
Dim Width As Long
Dim Height As Long

Dim LastHit As Boolean
Dim StartX As Long
Dim StartY As Long


'Create A region to shape the Form
Width = Pic.ScaleX(Pic.Width, vbTwips, vbPixels)
Height = Pic.ScaleY(Pic.Height, vbTwips, vbPixels)
'Create a new Region
rgnMain = CreateRectRgn(0, 0, Width, Height)
dcMain = CreateCompatibleDC(Pic.hDC)
'Get the picture we us for this calculation
bmpMain = SelectObject(dcMain, Pic.Picture.Handle)

'Move thru it
For Y = 0 To Height
For X = 0 To Width
RGBColor = GetPixel(dcMain, X, Y)
'Found a transparent spot
'make it also tramsparent on the region
If RGBColor = TransColor And LastHit = False Then
LastHit = True
StartX = X
StartY = Y
ElseIf LastHit = True And RGBColor <> TransColor Then
LastHit = False
'we found Transparent Pixels now create a region
If Y > StartY Then 'We found more than one row of transparent pixels
If StartX > 0 Then 'We didnt start at point 0 so create the first line
rgnPixel = CreateRectRgn(StartX, StartY, Width + 1, StartY + 1) 'The first line from start to the end
CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
DeleteObject rgnPixel
Else
StartY = StartY - 1 'Tell the code to do one line more
End If
If Y > StartY + 1 Then
rgnPixel = CreateRectRgn(0, StartY + 1, Width + 1, Y) 'Now line 2 to y
CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
DeleteObject rgnPixel
End If
rgnPixel = CreateRectRgn(0, Y, X, Y + 1) 'the last line (x because the actual pixel is not ok)
CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
DeleteObject rgnPixel
Else 'We are still in the same line so create only the pixels we found
rgnPixel = CreateRectRgn(StartX, Y, X, Y + 1)
CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
DeleteObject rgnPixel
End If
End If
Next X
Next Y

'Remove unused
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain

'Get the Region Data so we can store it later
If rgnMain <> 0 Then
ByteCtr = GetRegionData(rgnMain, 0, ByVal 0&)
If ByteCtr > 0 Then
ReDim RgnData(0 To ByteCtr - 1)
ByteCtr = GetRegionData(rgnMain, ByteCtr, RgnData(0))
End If
'Shape the form
SetWindowRgn Pic.hWnd, rgnMain, True
End If
CalculationDone = True
Pic.AutoSize = True
End Sub

