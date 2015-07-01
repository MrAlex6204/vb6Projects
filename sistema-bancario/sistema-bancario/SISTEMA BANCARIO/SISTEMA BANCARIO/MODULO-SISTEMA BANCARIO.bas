Attribute VB_Name = "Module1"
Public CN As New Connection
Public RS As New Recordset
Public Sub BD()
CN.Provider = "MICROSOFT.JET.OLEDB.4.0"
CN.Open App.Path + "\SISTEMA BANCARIO.MDB"
End Sub

Public Sub MOVER(R As Recordset, X As Integer)
Select Case X
       Case 0: R.MoveFirst
       Case 1: R.MovePrevious
       If R.BOF Then R.MoveFirst
       Case 2: R.MoveNext
       If R.EOF Then R.MoveLast
       Case 3: R.MoveLast
End Select
End Sub

Public Sub LIMPIAR(F As Form)
Dim C As Control
For Each C In F
If TypeOf C Is TextBox Or TypeOf C Is ComboBox Then C = ""
If TypeOf C Is OptionButton Then C.Value = False
If TypeOf C Is CheckBox Then C.Value = 0
If TypeOf C Is ListBox Then C.Clear
Next
End Sub

Public Function BUSCAR(T As String, CAMPO As String, DATO As String, NCOL As Integer) As String
Dim RSTEMP As New Recordset
SQL = "SELECT*FROM " + T + " WHERE " + CAMPO + "='" + DATO + "'"
Set RSTEMP = CN.Execute(SQL)
If RSTEMP.EOF Then
   BUSCAR = " "
   Else
   BUSCAR = RSTEMP(NCOL)
End If
RSTEMP.Close
End Function

Public Sub COMBO(T As String, C As ComboBox, NCOL As Integer)
Dim SQLC As String
SQLC = "SELECT*FROM " + T
RS.Open SQLC, CN, adOpenDynamic, adLockOptimistic
Do While Not RS.EOF
   C.AddItem RS(NCOL)
   RS.MoveNext
Loop
   RS.Close
End Sub

'Esta función permite hallar la suma de una determinada
'columna dentro de un control Microsoft Hierarchical
'Flexgrid.
Public Function SUMACOLUMNAMHF(NUMEROCOLUMNA As Integer, H As MSHFlexGrid) As Single
Dim ST As Single, I As Integer
ST = 0
For I = 0 To H.Rows - 1
  If Not IsNull(H.TextMatrix(I, NUMEROCOLUMNA)) And IsNumeric(H.TextMatrix(I, NUMEROCOLUMNA)) Then
    ST = ST + H.TextMatrix(I, NUMEROCOLUMNA)
  End If
Next
SUMACOLUMNAMHF = ST
End Function

'Esta función se usa en caso de que el dato a buscar
'se encuentre en un campo del tipo númerico
Public Function BUSCARDATONUM(T As String, CAMPO As String, DATO As String, NCOL As Integer) As String
Dim RSTEMP As New Recordset
SQL = "SELECT*FROM " + T + " WHERE " + CAMPO + "= " + DATO + " "
Set RSTEMP = CN.Execute(SQL)
If RSTEMP.EOF Then
   BUSCARDATONUM = " "
   Else
   BUSCARDATONUM = RSTEMP(NCOL)
End If
RSTEMP.Close
End Function

