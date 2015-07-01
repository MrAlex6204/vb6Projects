Attribute VB_Name = "modFormManagement"
Option Explicit

Public Sub UpdateWindowList(Window As Variant)
    Dim lngi As Long
    Dim lngj As Long
    
    For lngi = 0 To UBound(Window) - 1
        If Window(lngi).Visible = False Then
            For lngj = lngi To UBound(Window) - 1
                Set Window(lngj) = Window(lngj + 1)
            Next
            ReDim Preserve Window(UBound(Window) - 1)
            UpdateWindowList Window
            Exit Sub
        End If
    Next
End Sub

Public Sub KillWindows(Window As Variant)
    Dim lngi As Long
    
    For lngi = 0 To UBound(Window) - 1
        Unload Window(lngi)
    Next
    ReDim Window(0)
End Sub

