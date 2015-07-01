Attribute VB_Name = "Datos"
Public BDD As Database 'Objeto para manejar la base de datos
Public TBL As Recordset  'Objeto para manejar la Tabla
Public SQL As String
   
   Public NomCajero As String
   Public sItemData As String
   Public strData As String 'Variable Que Sirvira Para Localizar Un Dato
   Public strConnect As String
   Public BaseDatosOpen As New ADODB.Recordset
   
   Public PathConection As String
   Public articulos As Integer
   Public total As Double
   Public Dia As Date
   Public tipCambio As String
   Public Encontrado As Boolean


Sub AbrirBase(SqlCommando As String)
BaseDatosOpen.Open SqlCommando, strConnect, adOpenKeyset, adLockOptimistic
End Sub
Sub CerrarBase()
BaseDatosOpen.Close
End Sub
Sub SetBase()
Set BDD = OpenDatabase(App.Path & "\Tienda.mdb")
End Sub
Sub ConectBase()

strPath = App.Path & "\Tienda.mdb"

strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      
End Sub
Sub Buscar(ByVal Buscar As String, ByVal Campo As String)
Encontrado = True
strData = Campo + " = '" & Buscar & "'"
'El RS.Find strData sirve Para EContrar Lo que hay en la variable strData

    BaseDatosOpen.Find strData
    
    If BaseDatosOpen.EOF = True Then
        If MsgBox("NO SE ENCONTRO ARTICULO !!! ¿DESEA AGREGARLO ?", vbYesNo, "ALTA PRODUCTO") = 6 Then
        
        End If
        
        Encontrado = False
        Exit Sub
        Exit Sub
    End If
End Sub
