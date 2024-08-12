VERSION 5.00
Begin VB.Form FormUserProfile 
   Caption         =   "UserProfile"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   16050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProfile 
      Caption         =   "Profile"
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.CommandButton cmdToHome 
         Caption         =   "Regresar a Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   7560
         TabIndex        =   1
         Top             =   6240
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "Eliminar usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   7560
         TabIndex        =   8
         Top             =   3600
         Width           =   2600
      End
      Begin VB.CommandButton cmdUpdateUser 
         Caption         =   "Modificar Información"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   7560
         TabIndex        =   9
         Top             =   4680
         Width           =   2600
      End
      Begin VB.TextBox textName 
         Height          =   450
         Left            =   2040
         TabIndex        =   2
         Top             =   1560
         Width           =   3800
      End
      Begin VB.TextBox textLastName 
         Height          =   450
         Left            =   2040
         TabIndex        =   3
         Top             =   2430
         Width           =   3800
      End
      Begin VB.TextBox textEmail 
         Height          =   450
         Left            =   2040
         TabIndex        =   4
         Top             =   3300
         Width           =   3800
      End
      Begin VB.TextBox textPassword 
         Height          =   450
         Left            =   2040
         TabIndex        =   5
         Top             =   4170
         Width           =   3800
      End
      Begin VB.TextBox textAge 
         Height          =   450
         Left            =   2040
         TabIndex        =   6
         Top             =   5040
         Width           =   3800
      End
      Begin VB.CommandButton cmdSignup 
         Caption         =   "Cerrar Sesión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2520
         TabIndex        =   7
         Top             =   6240
         Width           =   2600
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
         Left            =   2160
         TabIndex        =   14
         Top             =   4680
         Width           =   3000
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
         Left            =   2160
         TabIndex        =   13
         Top             =   3810
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
         Left            =   2160
         TabIndex        =   12
         Top             =   2940
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
         Left            =   2160
         TabIndex        =   11
         Top             =   2070
         Width           =   1455
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
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Image userPhoto 
         Height          =   2535
         Left            =   6960
         Top             =   720
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FormUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSignup_Click()
    FormLogin.Show
    FormUserProfile.Hide
End Sub


Private Sub cmdToHome_Click()
    FormMain.Show
    FormUserProfile.Hide
End Sub

Private Sub cmdDeleteUser_Click()
    
    Dim deleteConfirmation As Integer
    deleteConfirmation = MsgBox("¿Deseas eliminar tu usuario?", vbOKCancel, "Confirmación")

    If deleteConfirmation = vbOK Then
    
        
        Dim sqlQuery As String
        sqlQuery = "delete from users where id_user =" & user.Id
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        FormLogin.Show
        FormUserProfile.Hide
        MsgBox "Usuario eliminado."
        
    Else
        MsgBox "Operación cancelada."
    End If
    
End Sub

Private Sub cmdUpdateUser_Click()
        Dim deleteConfirmation As Integer
    deleteConfirmation = MsgBox("¿Deseas modificar los datos?", vbOKCancel, "Confirmación")

    If deleteConfirmation = vbOK Then
    
        Dim name As String
        name = textName.Text
        Dim lastName As String
        lastName = textLastName.Text
        Dim email As String
        email = textEmail.Text
        Dim pass As String
        pass = textPassword.Text
        Dim age As String
        age = textAge.Text
        
        Dim sqlQuery As String
        sqlQuery = "delete from users where id_user =" & user.Id
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        FormLogin.Show
        FormUserProfile.Hide
        MsgBox "Usuario eliminado."
        
    Else
        MsgBox "Operación cancelada."
    End If
End Sub

Private Sub Form_Load()
    cmdUpdateUser.Enabled = False
    
End Sub

Private Sub textAge_Change()
    If textAge.Text <> user.age Then
        cmdUpdateUser.Enabled = True
    Else
        cmdUpdateUser.Enabled = False
    End If
End Sub

Private Sub textEmail_Change()
    If textEmail.Text <> user.email Then
        cmdUpdateUser.Enabled = True
    Else
        cmdUpdateUser.Enabled = False
    End If
End Sub

Private Sub textLastName_Change()

    If textLastName.Text <> user.lastName Then
        cmdUpdateUser.Enabled = True
    Else
        cmdUpdateUser.Enabled = False
    End If
    
End Sub

Private Sub textName_Change()
    
    If textName.Text <> user.name Then
        cmdUpdateUser.Enabled = True
    Else
        cmdUpdateUser.Enabled = False
    End If

End Sub

Private Sub textPassword_Change()
    If textPassword.Text <> user.Password Then
        cmdUpdateUser.Enabled = True
    Else
        cmdUpdateUser.Enabled = False
    End If
End Sub


Function InfoModified()

End Function
