Attribute VB_Name = "Efects"

 Option Explicit
 '*************************************
 'Copyright © 2001 by Alexander Anikin
 '& DREAM Interactive - QQQ
 'e-mail: aka@i.com.ua
 'For more my code samples visit:
 'http://www.i.com.ua/~aka
 '*************************************
 Public Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean 'only Windows 98 or Latter
 Dim Num As Byte, nN%, nBlend&

Sub Run_Blending(Picture1 As PictureBox, Picture2 As PictureBox)
 Num = CByte(255)
 nN = 5
Do
 DoEvents
 '***********************************************
  nBlend = vbBlue - CLng(Num) * (vbYellow + 1)
 'It's Magic Formula is
 'Alchemical Mixture of Elements of Gold & Sky
 'It's obtained by an almost mystical way
 '***********************************************
 Num = CByte(Num) - nN
 If Num = CByte(0) Then
   nN = -5
 ElseIf Num = CByte(255) Then
   nN = 5
 End If
Picture1.Cls
 AlphaBlend Picture1.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, nBlend
Loop
End Sub
Sub BlendingForm(picSrc As PictureBox, FOrm1 As Form)
 Num = CByte(255)
 nN = 5
Do
 DoEvents
 '***********************************************
  nBlend = vbBlue - CLng(Num) * (vbYellow + 1)
 'It's Magic Formula is
 'Alchemical Mixture of Elements of Gold & Sky
 'It's obtained by an almost mystical way
 '***********************************************
 Num = CByte(Num) - nN
 If Num = CByte(0) Then
   nN = -5
 ElseIf Num = CByte(255) Then
   nN = 5
 End If
FOrm1.Cls
 AlphaBlend FOrm1.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hDC, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, nBlend
Loop
End Sub
