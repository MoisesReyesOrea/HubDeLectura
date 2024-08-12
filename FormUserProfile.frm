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
   Begin VB.Frame frmSignup 
      Caption         =   "Signup"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdToHome 
         Caption         =   "Regresar"
         Height          =   700
         Left            =   7560
         TabIndex        =   1
         Top             =   6240
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "Eliminar usuario"
         Height          =   700
         Left            =   7560
         TabIndex        =   8
         Top             =   3600
         Width           =   2600
      End
      Begin VB.CommandButton cmdUpdateUser 
         Caption         =   "Modificar Información"
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
         Height          =   2175
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



Private Sub lblToHome_Click()
    FormMain.Show
    FormUserProfile.Hide
End Sub

