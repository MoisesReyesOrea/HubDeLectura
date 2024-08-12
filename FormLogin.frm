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
      Height          =   7695
      Left            =   9120
      TabIndex        =   13
      Top             =   600
      Width           =   8295
      Begin VB.TextBox textAge 
         Height          =   450
         Left            =   2400
         TabIndex        =   10
         Top             =   5280
         Width           =   3800
      End
      Begin VB.CommandButton cmdSignup 
         Caption         =   "Registrarse"
         Height          =   700
         Left            =   3000
         TabIndex        =   11
         Top             =   6000
         Width           =   2600
      End
      Begin VB.TextBox textPassConf 
         Height          =   450
         Left            =   2400
         TabIndex        =   9
         Top             =   4320
         Width           =   3800
      End
      Begin VB.TextBox textPassword 
         Height          =   450
         Left            =   2400
         TabIndex        =   8
         Top             =   3405
         Width           =   3800
      End
      Begin VB.TextBox textEmail 
         Height          =   450
         Left            =   2400
         TabIndex        =   7
         Top             =   2475
         Width           =   3800
      End
      Begin VB.TextBox textLastName 
         Height          =   450
         Left            =   2400
         TabIndex        =   6
         Top             =   1530
         Width           =   3800
      End
      Begin VB.TextBox textName 
         Height          =   450
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   3800
      End
      Begin VB.Label Label1 
         Caption         =   "Confirma contraseña"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   3990
         Width           =   3000
      End
      Begin VB.Label lblName 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLastName 
         Caption         =   "Apellido"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label lblEmail 
         Caption         =   "Correo"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label lblPassword 
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   3045
         Width           =   1455
      End
      Begin VB.Label lblAge 
         Caption         =   "Edad"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label LblToLogin 
         Caption         =   "Ya tengo cuenta"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   6960
         Width           =   1695
      End
   End
   Begin VB.Frame frmLogin 
      Caption         =   "Login"
      Height          =   7695
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      Begin VB.TextBox txtLoginPassword 
         Height          =   450
         Left            =   2040
         TabIndex        =   2
         Top             =   3120
         Width           =   3800
      End
      Begin VB.TextBox txtLoginEmail 
         Height          =   450
         Left            =   2040
         TabIndex        =   1
         Top             =   2040
         Width           =   3800
      End
      Begin VB.CommandButton cmnLogin 
         Caption         =   "Iniciar sesión"
         Height          =   700
         Left            =   2520
         TabIndex        =   3
         Top             =   4080
         Width           =   2600
      End
      Begin VB.CommandButton cmdToSignup 
         Caption         =   "Crear Cuenta"
         Height          =   700
         Left            =   2520
         TabIndex        =   4
         Top             =   5280
         Width           =   2600
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Correo"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   1680
         Width           =   1455
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
    
    ' Condicion para validar datos ingresados
    If Len(Trim(loginEmail)) > 0 And Len(Trim(loginPass)) > 0 Then
        
    Dim querySql As String
    querySql = "select * from users where email = '" + loginEmail + "'"
    
    cn.Open connectionData
    rs.Source = "users" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open querySql, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    ' Si el email no se encuentra entonces
    If rs.BOF And rs.EOF Then
        MsgBox "No se encontro registro del usuario ingresado", , Login
    Else
        ' Aquí el Recordset tiene al menos un registro
        If Not loginPass = rs.Fields("password").value Then
        MsgBox "La contraseña no coincide.", , Login
        
        Else
            MsgBox "Inicio de sesión exitoso."
            rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
            'Debug.Print rs.Fields("name")
    
            user.Id = rs.Fields("id_user")
            user.Name = rs.Fields("name")
            user.LastName = rs.Fields("last_name")
            user.Email = rs.Fields("email")
            user.Password = rs.Fields("password")
            user.Age = rs.Fields("age")
    
            Debug.Print user.Id, user.Name, user.LastName, user.Email, user.Password, user.Age
        
            FormUserProfile.textName.Text = user.Name
            FormUserProfile.textLastName.Text = user.LastName
            FormUserProfile.textEmail.Text = user.Email
            FormUserProfile.textPassword.Text = user.Password
            FormUserProfile.textAge.Text = user.Age
    
            closeConnectionDB
    
            FormMain.Show
            FormLogin.Hide
        
        End If
        
    End If
    
    Else
        MsgBox "Llena los campos requeridos", , Login
    End If
    
    If rs.State = adStateOpen Then
        closeConnectionDB
    End If
    
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
    
    'Abrimos la base de datos
    cn.Open connectionData
    
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
    frmLogin.Visible = True
    frmSignup.Visible = False
    
End Sub


Private Sub LblToLogin_Click()
    frmLogin.Visible = True
    frmSignup.Visible = False
End Sub
