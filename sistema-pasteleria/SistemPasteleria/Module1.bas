Attribute VB_Name = "modulosistema"
Global cn As ADODB.Connection
Global rspro As ADODB.Recordset
Global rscli As ADODB.Recordset
Global rslin As ADODB.Recordset
Global RsTemporal As ADODB.Recordset
Global rsest As ADODB.Recordset
Global rsfor As ADODB.Recordset
Global rsfac As ADODB.Recordset
Global rsdetfac As ADODB.Recordset
Global rscar As ADODB.Recordset
Global rsper As ADODB.Recordset
Global rsreg As ADODB.Recordset
Global rsdis As ADODB.Recordset
Global rsmar As ADODB.Recordset
Global rsusu As ADODB.Recordset
Global rsprov As ADODB.Recordset
Global rsgra As ADODB.Recordset
Global rsmed As ADODB.Recordset
Global rsguiai As ADODB.Recordset
Global rsdguiai As ADODB.Recordset
Global rsguia As ADODB.Recordset
Global rsdguia As ADODB.Recordset
Global rsbol As ADODB.Recordset
Global rsdbol As ADODB.Recordset
Global rspedint As ADODB.Recordset
Global rsdpedido As ADODB.Recordset
Global rstmp As ADODB.Recordset
Global rspago As ADODB.Recordset
Global rsorden As ADODB.Recordset
Global rsdorden As ADODB.Recordset
Global rstrans As ADODB.Recordset
Global rsve As ADODB.Recordset
Global rsdo As ADODB.Recordset
Global rsddo As ADODB.Recordset
Global sqlcli As String
Global sqllin As String
Global sqlpro As String
Global sqlTemporal As String
Global sqlest As String
Global sqlfor As String
Global sqlfac As String
Global sqldetfac As String
Global sqlcar As String
Global sqlper As String
Global sqlreg As String
Global sqldis As String
Global sqlmar As String
Global sqlusu As String
Global sqlprov As String
Global sqlgra As String
Global sqlmed As String
Global sqlguiai As String
Global sqlguia As String
Global sqldguia As String
Global sqlbol As String
Global sqldbol As String
Global sqlpedint As String
Global sqldpedido As String
Global sqltmp As String
Global sqlpago As String
Global sqldguiai As String
Global sqlorden As String
Global sqldorden As String
Global sqltrans As String
Global sqlve As String
Global sqldo As String
Global sqlddo As String


Public Sub conectar()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.ConnectionString = App.Path + "\base.mdb"
cn.Open
If cn.State = stateclosed Then
MsgBox ("CONEXION INCORRECTA")
End
Else
MsgBox ("CONEXION CORRECTA")
End If
End Sub
Public Sub desconectar()
cn.Close
Set cn = Nothing
MsgBox ("CONEXION CERRADA")
End Sub

Public Sub activacli()
Set rscli = New ADODB.Recordset
sqlcli = "select *from clientes"
rscli.Open sqlcli, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activalin()
Set rslin = New ADODB.Recordset
sqllin = "select *from lineas"
rslin.Open sqllin, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activapro()
Set rspro = New ADODB.Recordset
sqlpro = "select *from productos"
rspro.Open sqlpro, cn, adOpenStatic, adLockOptimistic

End Sub

Public Sub activaest()
Set rsest = New ADODB.Recordset
sqlest = "select *from estados"
rsest.Open sqlest, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activafor()
Set rsfor = New ADODB.Recordset
sqlfor = "select *from formapago"
rsfor.Open sqlfor, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activafac()
Set rsfac = New ADODB.Recordset
sqlfac = "select *from facturas"
rsfac.Open sqlfac, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadetfac()
Set rsdetfac = New ADODB.Recordset
sqldetfac = "select *from detallefacturas"
rsdetfac.Open sqldetfac, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activacar()
Set rscar = New ADODB.Recordset
sqlcar = "select *from cargos"
rscar.Open sqlcar, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activaper()
Set rsper = New ADODB.Recordset
sqlper = "select *from personal"
rsper.Open sqlper, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activareg()
Set rsreg = New ADODB.Recordset
sqlreg = "select *from registros"
rsreg.Open sqlreg, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadis()
Set rsdis = New ADODB.Recordset
sqldis = "select *from distritos"
rsdis.Open sqldis, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activamar()
Set rsmar = New ADODB.Recordset
sqlmar = "select *from marcas"
rsmar.Open sqlmar, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activausu()
Set rsusu = New ADODB.Recordset
sqlusu = "select *from usuarios"
rsusu.Open sqlusu, cn, adOpenStatic, adLockOptimistic
End Sub


Public Sub activaprov()
Set rsprov = New ADODB.Recordset
sqlprov = "select *from proveedores"
rsprov.Open sqlprov, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activagra()
Set rsgra = New ADODB.Recordset
rsgra.CursorLocation = adUseClient
sqlgra = "Select fec_emi,(total) from facturas" '& "Group by Month (fec_emi)"
rsgra.Open sqlgra, cn, adOpenStatic, adLockReadOnly
rsgra.ActiveConnection = cn
End Sub
Public Sub activamed()
Set rsmed = New ADODB.Recordset
sqlmed = "select *from unidadmedida"
rsmed.Open sqlmed, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activaguiai()
Set rsguiai = New ADODB.Recordset
sqlguiai = "select *from guiainterna"
rsguiai.Open sqlguiai, cn, adOpenStatic, adLockOptimistic
End Sub


Public Sub activaguia()
Set rsguia = New ADODB.Recordset
sqlguia = "select *from guiaremision"
rsguia.Open sqlguia, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadguia()
Set rsdguia = New ADODB.Recordset
sqldguia = "select *from detalleguiaremision"
rsdguia.Open sqldguia, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activabol()
Set rsbol = New ADODB.Recordset
sqlbol = "select *from boletas"
rsbol.Open sqlbol, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadbol()
Set rsdbol = New ADODB.Recordset
sqldbol = "select *from detalleboletas"
rsdbol.Open sqldbol, cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub activapedidoint()
Set rspedint = New ADODB.Recordset
sqlpedint = "select *from pedidointerno"
rspedint.Open sqlpedint, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadpedido()
Set rsdpedido = New ADODB.Recordset
sqldpedido = "select *from detallepedidointerno"
rsdpedido.Open sqldpedido, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activapago()
Set rspago = New ADODB.Recordset
sqlpago = "select *from pagos"
rspago.Open sqlpago, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadguiai()
Set rsdguiai = New ADODB.Recordset
sqldguiai = "select *from detalleguiainterna"
rsdguiai.Open sqldguiai, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activaordenc()
Set rsorden = New ADODB.Recordset
sqlorden = "select *from ordencompra"
rsorden.Open sqlorden, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activadorden()
Set rsdorden = New ADODB.Recordset
sqldorden = "select *from detalleordencompra"
rsdorden.Open sqldorden, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activatrans()
Set rstrans = New ADODB.Recordset
sqltrans = "select *from transportistas"
rstrans.Open sqltrans, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activave()
Set rsve = New ADODB.Recordset
sqlve = "select *from vehiculos"
rsve.Open sqlve, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activado()
Set rsdo = New ADODB.Recordset
sqldo = "select *from documentoentrada"
rsdo.Open sqldo, cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub activaddo()
Set rsddo = New ADODB.Recordset
sqlddo = "select *from detalledocumentoentrada"
rsddo.Open sqlddo, cn, adOpenStatic, adLockOptimistic
End Sub
