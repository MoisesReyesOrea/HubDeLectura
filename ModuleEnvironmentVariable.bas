Attribute VB_Name = "ModuleEnvironmentVariable"
' Variables para conectar a DB SQL Server
Public Const providerDB As String = "SQLOLEDB"
Public Const sourceDB As String = "LAPTOPS1"
Public Const nameDB As String = "HubDeLectura"
Public Const userIdDB As String = "usersql"
Public Const passDB As String = "root"

Public Const connectionData As String = "Provider=" + providerDB + ";Data Source=" + sourceDB + ";Initial Catalog=" + nameDB + ";User ID=" + userIdDB + ";Password=" + passDB + ";"

