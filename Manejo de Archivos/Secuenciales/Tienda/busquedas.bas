Attribute VB_Name = "busquedas"
Public NumCajero, Nombre As String
Public Precio, Descrip, NumArt As String
Public Archivo As Integer
Public CajeroOnline As String
Public Total As Double
Public Index As String
Public Index2 As Integer
Public Existe As Boolean

Function BuscarArt(BuscArticulo As String) As Boolean
Archivo = FreeFile()
Open App.Path + "\Articulos.dat" For Input As #Archivo
Do While Not EOF(Archivo)
    Input #Archivo, NumArt, Precio, Descrip
        If BuscArticulo = NumArt Then
        BuscarArt = True
    Exit Do
Else
BuscarArt = False
End If
Loop
Close #Archivo

End Function

Function BuscarCajero(Cajero As String) As Boolean
Archivo = FreeFile()
Open App.Path + "\Cajeros.dat" For Input As #Archivo
Do While Not EOF(Archivo)
    Input #Archivo, NumCajero, Nombre
    If Cajero = NumCajero Then
        BuscarCajero = True
    Exit Do
Else
BuscarCajero = False
End If
Loop
Close #Archivo

End Function
Function Eliminar(Elim As String, file As String) As Boolean
'Existe Toma un Valor de False por inicio
'Si si valor Cambia kiere decir q si se encontro el articulo
'de lo contrario su valor keda intacto
Existe = False
On Error Resume Next
Dim ArchTemp As Integer

Archivo = FreeFile()
ArchTemp = 19

Open file For Input As #Archivo
Open App.Path + "\Temp.bkup" For Output As #ArchTemp

Do While Not EOF(Archivo)

    Input #Archivo, Index, NumArt, Descrip, Precio
    If Elim <> Index Then
    'Si es diferente lo graba al archivo Temporal
    Print #ArchTemp, Index
    Print #ArchTemp, NumArt
    Print #ArchTemp, Descrip
    Print #ArchTemp, Precio
    Else
    'De lo Contrario No graba el articulo y le
    'Indicamos al la variable existe que no grabe y q continue con el ciclo
    Existe = True
    End If
    
Loop
Close #Archivo
Close #ArchTemp
'Actualiza el archivo

Open file For Output As #ArchTemp
Open App.Path + "\Temp.bkup" For Input As #Archivo

Do While Not EOF(Archivo)

    Input #Archivo, Index, NumArt, Descrip, Precio
  
    Print #ArchTemp, Index
    Print #ArchTemp, NumArt
    Print #ArchTemp, Descrip
    Print #ArchTemp, Precio
   
  
    
Loop
Close #Archivo
Close #ArchTemp
Kill App.Path + "\Temp.bkup"
End Function
