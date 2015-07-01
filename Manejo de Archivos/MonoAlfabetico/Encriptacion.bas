Attribute VB_Name = "Encriptacion"
Public Sub Generar(Clave As String)
'////////////Generar Clave////////////////////////
Dim X, N, i, j As Integer
Dim Linea, Caracter As String
Dim Encontrado As Boolean
Encontrado = False

VeraSoft.Progress.Visible = True
VeraSoft.Progress.Max = 27
VeraSoft.Progress.Min = 0


X = Len(Clave)

Grabar (Clave + "*") 'Graba la inicio del Archivo
'La clave y pone un * al final de grabar

N = 0
Linea = ""
Do While (27 > N)
 
        VeraSoft.Progress.Value = N
        j = 0
        
        Do While (X > j)
            
            'Extrae Una Letra de La Clave
            'En cada Ciclo
            Caracter = Mid(Clave, j + 1, 1)
            
            'Si Coinciden la Letra Extraida Se Sale del While
            'Para Obtener la Siguiente Letra del Arreglo
            'a comparar
            If UCase(Caracter) = UCase(ABC(N)) Then
            Encontrado = True
            Exit Do
            End If
            j = j + 1
            
       Loop
       
        'Si en La Busqueda de la Letra Extraida
        'Dio Falso La Suma ala Variable Linea
        If Encontrado = False Then
        Linea = Linea + ABC(N)
        End If
        'Se recetea la Variable en dado Caso de que
        'Cambio a True
        Encontrado = False
         
         'Si la Variable Acumuladora Linea es Igual
         'al numero de letras que tiene la Parabra CLave o
         'N ya esta en la ultima Letra del ABCDario
         'Graba en el Archivo su Contenido
        If (Len(Linea)) = X Or N = 26 Then
        If N = 26 Then
        Grabar (Linea + "**************************")
        Else
        Grabar (Linea + "*")
        End If
        'Se Recea la Variable Linea
        'Para Proseguir a Escribir en el Archivo en la sig. Fila
        Linea = ""
        End If
    N = N + 1

Loop

VeraSoft.Progress.Value = N


End Sub

Public Sub GenerarEncriptacioFiles(Longitud As String)
Dim Contenido, Chr, CadenaTotal, Cadena1, Cadena2, VeriFic As String
Dim Posx, LongCadena, PosAbc As Integer

Open App.path + "\Encript.Dat" For Input As #20
Open App.path + "\Encript2.Dat" For Output As #10

LongCadena = Len(Longitud)

PosAbc = 0
Chr = ""

VeraSoft.Progress.Visible = True
VeraSoft.Progress.Max = 27
VeraSoft.Progress.Min = 0


'Acomoda el AbcDario en La Primera Columna del Archivo Encript
Do While Not EOF(20)
                
        VeraSoft.Progress.Value = PosAbc
        
        Line Input #20, Contenido
        LongCadena = Len(Contenido)
        Cadena = Mid(Contenido, 2, (LongCadena - 1))
        Chr = ABC(PosAbc) + Cadena
        Print #10, UCase(Chr)
        PosAbc = PosAbc + 1
       
  
Loop
Close #20
Close #10



'Continua Acomodando el AbcDario En las Siguientes Columnas hasta
'Llegar al fin del AbcDario



Open App.path + "\Encript2.Dat" For Input As #40
Open App.path + "\Encript3.Dat" For Output As #50
VeriFic = ""
Posx = 2

'Bucle hasta que Llege al Fin del AbcDario
Do While (PosAbc < 25)

VeraSoft.Progress.Value = PosAbc

        Do While Not EOF(40)
        Line Input #40, Contenido
        
        VeriFic = Mid(Contenido, Posx, 1)
        
      
        
        If VeriFic <> "*" Then
            Cadena1 = Mid(Contenido, 1, Posx - 1)
            Cadena2 = Mid(Contenido, Posx + 1, (LongCadena + 1) - Posx)
            CadenaTotal = Cadena1 + ABC(PosAbc) + Cadena2
            Print #50, CadenaTotal
            PosAbc = PosAbc + 1
            VeraSoft.Progress.Value = PosAbc
        Else
            'Recetea la Variable
            VeriFic = ""
            Print #50, Contenido
            Exit Do
        End If
        
        Loop
        
        Close #40
        Close #50
            
          
        
        'Actualizacion del Archivo
        Kill App.path + "\Encript2.Dat"
        'Renombra el Nuevo por el el Viejo
        FileCopy App.path + "\Encript3.Dat", App.path + "\Encript2.Dat"
        'Elimina el Nuevo
        Kill App.path + "\Encript3.Dat"
        
        'Vuelve Abrir Los Archivos Pero ya Actualizados
        Open App.path + "\Encript2.Dat" For Input As #40
        Open App.path + "\Encript3.Dat" For Output As #50
        
        Posx = Posx + 1
        
Loop

Close #40
Close #50
Kill App.path + "\Encript3.Dat"


VeraSoft.Progress.Value = PosAbc
End Sub
Sub Encriptar(Texto As String)

Dim Contenido1, Contenido2, Chr, Chr2, ChrBuscar, Resultado As String
Dim Posx, LongCadena, LongCadena2, Lectura1, Lectura2, i, N As Integer
Lectura1 = 80
Lectura2 = 70

VeraSoft.Progress.Visible = True
VeraSoft.Progress.Min = 0

LongCadena = Len(Texto)
VeraSoft.Progress.Max = LongCadena


N = 0

Do While (LongCadena > N)

VeraSoft.Progress.Value = N

ChrBuscar = Mid(UCase(Texto), N + 1, 1)
Open "Encript2.Dat" For Input As #Lectura1
Open "Encript.Dat" For Input As #Lectura2
    Do While Not EOF(Lectura1)
    
    Line Input #Lectura1, Contenido1
    Line Input #Lectura2, Contenido2
    
    LongCadena2 = Len(Contenido1)
    
        i = 0
        
        Do While (LongCadena2 > i)
            Chr = Mid(Contenido1, i + 1, 1)
            Chr2 = Mid(Contenido2, i + 1, 1)
            If Chr = ChrBuscar Then
            EncriptacionText = EncriptacionText + Chr2
            Exit Do
        End If
        
        i = i + 1
        Loop
        
        If Chr = ChrBuscar Then
        Exit Do
        End If
        
        
        
    Loop
    Close #Lectura1
    Close #Lectura2


N = N + 1
Loop


VeraSoft.Progress.Value = N
End Sub
Sub Desencriptardor(Texto As String)

Dim Contenido1, Contenido2, Chr, Chr2, ChrBuscar, Resultado As String
Dim Posx, LongCadena, LongCadena2, Lectura1, Lectura2, i, N As Integer
Lectura1 = 80
Lectura2 = 70

VeraSoft.Progress.Visible = True
VeraSoft.Progress.Min = 0

LongCadena = Len(Texto)
VeraSoft.Progress.Max = LongCadena
N = 0




Do While (LongCadena > N)
VeraSoft.Progress.Value = N

ChrBuscar = Mid(UCase(Texto), N + 1, 1)

Open "Encript.Dat" For Input As #Lectura1
Open "Encript2.Dat" For Input As #Lectura2

    Do While Not EOF(Lectura1)
    
    Line Input #Lectura1, Contenido1
    Line Input #Lectura2, Contenido2
    LongCadena2 = Len(Contenido1)
    
        i = 0
        Do While (LongCadena2 > i)
        Chr = Mid(Contenido1, i + 1, 1)
        Chr2 = Mid(Contenido2, i + 1, 1)
        
      '  MsgBox Chr + "=" + Chr2
        If Chr = ChrBuscar Then
        DesencriptacionText = DesencriptacionText + Chr2
        
        Exit Do
        End If
        
        i = i + 1
        Loop
        
        If Chr = ChrBuscar Then
        Exit Do
        End If
        
        
        
    Loop
    Close #Lectura1
    Close #Lectura2


N = N + 1
Loop

VeraSoft.Progress.Value = N

End Sub

