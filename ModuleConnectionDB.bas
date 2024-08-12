Attribute VB_Name = "ModuleConnectionDB"
Public cn As New ADODB.Connection  'Creamos el objeto Connection.
Public rs As New ADODB.Recordset 'Creamos el objeto Recordset.
Public cmd As New ADODB.Command    ' Creamos el objeto command


Function GetConnectionDB()

End Function


Function closeConnectionDB()
    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
End Function
