VERSION 5.00
Begin VB.Form FormBook 
   Caption         =   "Book"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18660
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   18660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBackHomeFB 
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
      Left            =   14520
      TabIndex        =   9
      Top             =   8400
      Width           =   2600
   End
   Begin VB.CommandButton cmdAddFavoriteFB 
      Caption         =   "Agregar a favoritos"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   8400
      Width           =   2600
   End
   Begin VB.CommandButton cmdAddCompletedFB 
      Caption         =   "Marcar como Leido"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   8400
      Width           =   2600
   End
   Begin VB.CommandButton cmdAddReadingFB 
      Caption         =   "Agregar a Leyendo"
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
      Left            =   7800
      TabIndex        =   6
      Top             =   8400
      Width           =   2600
   End
   Begin VB.CommandButton cmdAddNoWishedFB 
      Caption         =   "Marcar como NO deseado"
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
      Left            =   11160
      TabIndex        =   5
      Top             =   8400
      Width           =   2600
   End
   Begin VB.Label lblDescriptionBook 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1560
      TabIndex        =   4
      Top             =   5160
      Width           =   15615
   End
   Begin VB.Label lblGenreBook 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   3840
      Width           =   3800
   End
   Begin VB.Label lblYearBook 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   3800
   End
   Begin VB.Label lblAuthorBook 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   3800
   End
   Begin VB.Label lblTitleBook 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   3800
   End
End
Attribute VB_Name = "FormBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdGoBackHomeFB_Click()
    FormBook.Hide
    FormMain.Show
End Sub

Private Sub cmdAddFavoriteFB_Click()
        
        Dim sqlQuery As String
        sqlQuery = queryAddFavorite(user.Id, bookProperties.id_book)
    
        Set cmd = New ADODB.Command 'Activamos el command
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        MsgBox "Libro añadido a favoritos", vbInformation
    
End Sub

Private Sub cmdAddCompletedFB_Click()
        
        Dim sqlQuery As String
        sqlQuery = queryAddCompleted(user.Id, bookProperties.id_book)
        
        ' Función insertar libro a tabla relacional
        AddBookToTable (sqlQuery)
        
        MsgBox "Libro añadido a Completados", vbInformation
        
End Sub


Private Sub cmdAddReadingFB_Click()

        Dim sqlQuery As String
        sqlQuery = queryAddReading(user.Id, bookProperties.id_book)
        
        ' Función insertar libro a tabla relacional
        AddBookToTable (sqlQuery)
        
        MsgBox "Libro añadido a lista de Leyendo", vbInformation

End Sub

Private Sub cmdAddNoWishedFB_Click()

        Dim sqlQuery As String
        sqlQuery = queryAddNowished(user.Id, bookProperties.id_book)
        
        ' Función insertar libro a tabla relacional
        AddBookToTable (sqlQuery)
        
        MsgBox "Libro añadido a lista de no deseados", vbInformation
End Sub


Function AddBookToTable(ByVal sqlQuery As String)

    Set cmd = New ADODB.Command 'Activamos el command
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close

End Function
