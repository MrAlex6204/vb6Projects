Attribute VB_Name = "Module1"
Public Login, Nombre, Pwd As String
Public Archivo As Integer
Public Numero, Salario As String
 Public Function BuscarLogin(user As String) As Boolean
Archivo = FreeFile()
Open App.Path + "\Accesos.txt" For Input As #Archivo

Do While Not EOF(Archivo)
      Line Input #Archivo, Login
      Line Input #Archivo, Nombre
      Line Input #Archivo, Pwd
       
        If user = Login Then
        MsgBox "Encontrado"
        BuscarLogin = True
        Exit Do
       Else
       BuscarLogin = False
      End If
Loop

Close #Archivo

End Function
