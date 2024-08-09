VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Login"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17895
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   17895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSignup 
      Caption         =   "Signup"
      Height          =   7575
      Left            =   9120
      TabIndex        =   1
      Top             =   960
      Width           =   8295
      Begin VB.CommandButton cmdSignup 
         Caption         =   "Registrarse"
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   5400
         Width           =   2535
      End
      Begin VB.TextBox textAge 
         Height          =   450
         Left            =   2160
         TabIndex        =   10
         Top             =   3840
         Width           =   3800
      End
      Begin VB.TextBox textPassword 
         Height          =   450
         Left            =   2160
         TabIndex        =   9
         Top             =   3120
         Width           =   3800
      End
      Begin VB.TextBox textEmail 
         Height          =   450
         Left            =   2160
         TabIndex        =   8
         Top             =   2280
         Width           =   3800
      End
      Begin VB.TextBox textLastName 
         Height          =   450
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   3800
      End
      Begin VB.TextBox textName 
         Height          =   450
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   3800
      End
      Begin VB.Label LblToLogin 
         Caption         =   "Ya tengo cuenta"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   6240
         Width           =   1695
      End
   End
   Begin VB.Frame frmLogin 
      Caption         =   "Login"
      Height          =   7455
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   7575
      Begin VB.TextBox txtLoginPassword 
         Height          =   600
         Left            =   2880
         TabIndex        =   5
         Top             =   2640
         Width           =   3500
      End
      Begin VB.TextBox txtLoginEmail 
         Height          =   600
         Left            =   2880
         TabIndex        =   4
         Top             =   1440
         Width           =   3500
      End
      Begin VB.CommandButton cmnLogin 
         Caption         =   "Iniciar sesión"
         Height          =   855
         Left            =   2640
         TabIndex        =   3
         Top             =   4200
         Width           =   3375
      End
      Begin VB.CommandButton cmdToSignup 
         Caption         =   "Crear Cuenta"
         Height          =   735
         Left            =   2520
         TabIndex        =   2
         Top             =   5760
         Width           =   3375
      End
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    frmLogin.Visible = True
    frmSignup.Visible = False
    
End Sub

Private Sub cmdToSignup_Click()
    frmLogin.Visible = False
    frmSignup.Visible = True
End Sub

Private Sub cmnLogin_Click()
    Dim loginEmail As String
    Dim loginPass As String
    
    loginEmail = txtLoginEmail
    loginPass = txtLoginPassword
    
    Dim querySql As String
    querySql = "select * from users where email = '" + loginEmail + "'"
    
    'MsgBox loginEmail & " " & loginPass
    
    'Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
    'Dim rs As New ADODB.Recordset 'Creamos el objeto Recordset.
    'Set rs = New ADODB.Recordset ' activamos el Recordset
    
    cn.Open "Provider=SQLOLEDB;Data Source=LAPTOPS1;Initial Catalog=HubDeLectura;User ID=usersql;Password=root;"
    rs.Source = "books" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open querySql, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    'rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    'Debug.Print rs.Fields("name")
    
    Dim user As New Users
    user.Id = rs.Fields("id_user")
    user.Name = rs.Fields("name")
    user.LastName = rs.Fields("last_name")
    user.Email = rs.Fields("email")
    user.Password = rs.Fields("password")
    user.Age = rs.Fields("age")
    
    Debug.Print user.Id, user.Name, user.LastName, user.Email, user.Password, user.Age
    
    
    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close
    
    'Debug.Print rs.Fields("name") aqui ya no es posible acceder a la info de rs porque ya se cerrro

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
    FormUserProfile.textName.Text = user.Name
    FormUserProfile.textLastName.Text = user.LastName
    FormUserProfile.textEmail.Text = user.Email
    FormUserProfile.textPassword.Text = user.Password
    FormUserProfile.textAge.Text = user.Age
    
    FormMain.Show
    FormLogin.Hide
End Sub

Private Sub cmdSignup_Click()
    Dim signupName As String
    Dim signupLastName As String
    Dim signupEmail As String
    Dim signupPass As String
    Dim aignupAge As String
    
    signupName = textName
    signupLastName = textLastName
    signupEmail = textEmail
    signupPass = textPassword
    signupAge = textAge
    
    'Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
    'Dim cmd As ADODB.Command    ' Creamos el objeto command
    'Set cmd = New ADODB.Command 'Activamos el command
    
    'Abrimos la base de datos
    cn.Open "Provider=SQLOLEDB;Data Source=LAPTOPS1;Initial Catalog=HubDeLectura;User ID=usersql;Password=root;" '"Provider=SQLOLEDB;" & "Data Source=LAPTOPS1"
    
    cmd.ActiveConnection = cn
    cmd.CommandText = "INSERT INTO users (name, last_name, email, password, age) VALUES (?, ?, ?, ?, ?)"

    cmd.Parameters.Append cmd.CreateParameter("name", adVarChar, adParamInput, 50, signupName)
    cmd.Parameters.Append cmd.CreateParameter("last_name", adVarChar, adParamInput, 50, signupLastName)
    cmd.Parameters.Append cmd.CreateParameter("email", adVarChar, adParamInput, 50, signupEmail)
    cmd.Parameters.Append cmd.CreateParameter("password", adVarChar, adParamInput, 30, signupPass)
    cmd.Parameters.Append cmd.CreateParameter("age", adSmallInt, adParamInput, , signupAge)
    cmd.Execute

    ' Cerrar la conexión
    cn.Close
    
    MsgBox "Usuario " & signupName & " añadido", , "Signup"
    
    Exit Sub
    
End Sub


Private Sub LblToLogin_Click()
    frmLogin.Visible = True
    frmSignup.Visible = False
End Sub
