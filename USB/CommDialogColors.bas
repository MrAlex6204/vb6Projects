Attribute VB_Name = "DialColores"
Public El_Color As Long
' Función Api CHOOSECOLOR
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" ( _
                                            pChoosecolor As CHOOSECOLOR) As Long
                                            
' Estructura CHOOSECOLOR para configurar el cuadro de diálogo
Public Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Function Abrir_CommonDialog_Color(Formu As Form) As Long
    ' Array de tipo Byte dinámico
    Dim CustomColors() As Byte
    ' Variable para utilizar la estructura
    Dim cc As CHOOSECOLOR
    'array de tipo Long
    Dim Custcolor(16) As Long
    'Variable de retorno
    Dim lReturn As Long

    'Establecemos el tamaño de la extructura
    cc.lStructSize = Len(cc)
    'Le pasamos el hwnd del form a cc
    cc.hwndOwner = Formu.hWnd
    'Establecemos la instancia de nuestra aplicación a cc.Hinstance
    cc.hInstance = App.hInstance
    'Establecemos los colores convertidos a Unicode
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'El flag a 0 dialogo normal, en 2 dialogo completo
    cc.flags = 2
    
    'Mostramos el Cuadro de diálogo
    If CHOOSECOLOR(cc) <> 0 Then
        'Retornamos a nuestra función el valor elegido
        Abrir_CommonDialog_Color = cc.rgbResult
        'Para los colores personalizados
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        Abrir_CommonDialog_Color = -1
    End If
End Function


