VERSION 5.00
Begin VB.Form FormUserProfile 
   Caption         =   "UserProfile"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSignup 
      Caption         =   "Signup"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "Eliminar usuario"
         Height          =   735
         Left            =   7080
         TabIndex        =   8
         Top             =   4560
         Width           =   2655
      End
      Begin VB.CommandButton cmdUpdateUser 
         Caption         =   "Modificar Información"
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   5640
         Width           =   2535
      End
      Begin VB.TextBox textName 
         Height          =   450
         Left            =   2040
         TabIndex        =   6
         Top             =   1560
         Width           =   3800
      End
      Begin VB.TextBox textLastName 
         Height          =   450
         Left            =   2040
         TabIndex        =   5
         Top             =   2520
         Width           =   3800
      End
      Begin VB.TextBox textEmail 
         Height          =   450
         Left            =   2040
         TabIndex        =   4
         Top             =   3360
         Width           =   3800
      End
      Begin VB.TextBox textPassword 
         Height          =   450
         Left            =   2040
         TabIndex        =   3
         Top             =   4200
         Width           =   3800
      End
      Begin VB.TextBox textAge 
         Height          =   450
         Left            =   2040
         TabIndex        =   2
         Top             =   4920
         Width           =   3800
      End
      Begin VB.CommandButton cmdSignup 
         Caption         =   "Cerrar Sesión"
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   6480
         Width           =   2535
      End
      Begin VB.Label lblToHome 
         Alignment       =   2  'Center
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
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

