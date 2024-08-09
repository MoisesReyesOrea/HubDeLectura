Attribute VB_Name = "ModuleConnectionDB"
Public cn As New ADODB.Connection  'Creamos el objeto Connection.
Public rs As New ADODB.Recordset 'Creamos el objeto Recordset.
Public cmd As New ADODB.Command    ' Creamos el objeto command


Function GetConnectionDB()

Dim cn As New ADODB.Connection  'Creamos el objeto Connection.

Set rs = New ADODB.Recordset 'Activamos el Recordset
    'Abrimos la base de datos
cn.Open "Provider=SQLOLEDB;Data Source=LAPTOPS1;Initial Catalog=HubDeLectura;User ID=usersql;Password=root;" '"Provider=SQLOLEDB;" & "Data Source=LAPTOPS1"
rs.Source = "favorites" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
rs.Open "select * from favorites f JOIN books b on f.id_book = b.id_book where id_user = '1'", cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.

End Function


Function closeConnectionDB()
    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
End Function
