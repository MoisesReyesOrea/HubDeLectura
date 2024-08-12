VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMain 
   Caption         =   "FormMain"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20880
   LinkTopic       =   "Form1"
   ScaleHeight     =   11460
   ScaleWidth      =   20880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCompletedBooks 
      Caption         =   "CompletedBooks"
      Height          =   5175
      Left            =   4800
      TabIndex        =   27
      Top             =   3720
      Width           =   16815
      Begin VB.CommandButton Command6 
         Caption         =   "Ver libro"
         Height          =   700
         Left            =   480
         TabIndex        =   42
         Top             =   3960
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteCompleted 
         Caption         =   "Eliminar de libros completados"
         Height          =   1215
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   2775
      End
      Begin MSComctlLib.ListView ListCompletedBooks 
         Height          =   3255
         Left            =   3240
         TabIndex        =   28
         Top             =   1560
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame frmNoWished 
      Caption         =   "NoWished"
      Height          =   5055
      Left            =   3600
      TabIndex        =   31
      Top             =   6000
      Width           =   15255
      Begin VB.CommandButton Command7 
         Caption         =   "Ver libro"
         Height          =   700
         Left            =   360
         TabIndex        =   43
         Top             =   3240
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteNoWished 
         Caption         =   "Quitar de libros no deseados"
         Height          =   855
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListNoWished 
         Height          =   2655
         Left            =   3360
         TabIndex        =   32
         Top             =   1560
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Libros no deseados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   41
         Top             =   600
         Width           =   4695
      End
   End
   Begin VB.Frame fraFavorites 
      Caption         =   "Favorites"
      Height          =   3735
      Left            =   3120
      TabIndex        =   24
      Top             =   6480
      Width           =   16455
      Begin VB.CommandButton Command5 
         Caption         =   "Ver libro"
         Height          =   700
         Left            =   840
         TabIndex        =   39
         Top             =   2520
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteFavorite 
         Caption         =   "Eliminar de favoritos"
         Height          =   700
         Left            =   960
         TabIndex        =   29
         Top             =   600
         Width           =   2600
      End
      Begin MSComctlLib.ListView ListFavorites 
         Height          =   2295
         Left            =   4560
         TabIndex        =   25
         Top             =   1200
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Favoritos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   26
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.Frame fraHistory 
      Caption         =   "Reading"
      Height          =   4935
      Left            =   4320
      TabIndex        =   34
      Top             =   240
      Width           =   15495
      Begin VB.CommandButton Command8 
         Caption         =   "Ver libro"
         Height          =   700
         Left            =   1080
         TabIndex        =   44
         Top             =   3240
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteReading 
         Caption         =   "Eliminar de lista leyendo"
         Height          =   700
         Left            =   1200
         TabIndex        =   38
         Top             =   1080
         Width           =   2600
      End
      Begin MSComctlLib.ListView ListReadings 
         Height          =   2655
         Left            =   4440
         TabIndex        =   37
         Top             =   1680
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Libros que se están leyendo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   40
         Top             =   480
         Width           =   5055
      End
   End
   Begin VB.Frame fraNavbar 
      Caption         =   "Navbar"
      Height          =   11415
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin VB.CommandButton cmdPerfil 
         Caption         =   "Perfil"
         Height          =   1000
         Left            =   360
         TabIndex        =   6
         Top             =   9800
         Width           =   2175
      End
      Begin VB.CommandButton cmdNoWished 
         Caption         =   "No deseados"
         Height          =   1000
         Left            =   360
         TabIndex        =   5
         Top             =   6800
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadings 
         Caption         =   "Leyendo"
         Height          =   1000
         Left            =   360
         TabIndex        =   4
         Top             =   5300
         Width           =   2175
      End
      Begin VB.CommandButton cmdFavorites 
         Appearance      =   0  'Flat
         Caption         =   "Favoritos"
         Height          =   1000
         Left            =   360
         TabIndex        =   2
         Top             =   2300
         Width           =   2175
      End
      Begin VB.CommandButton cmdHome 
         Caption         =   "Inicio"
         Height          =   1000
         Left            =   360
         TabIndex        =   1
         Top             =   800
         Width           =   2175
      End
      Begin VB.CommandButton cmdCompletedBooks 
         Caption         =   "Completados"
         Height          =   1000
         Left            =   360
         TabIndex        =   3
         Top             =   3800
         Width           =   2175
      End
   End
   Begin VB.Frame fraHome 
      Caption         =   "Home"
      Height          =   10635
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   16575
      Begin VB.CommandButton cmdBookView 
         Caption         =   "Ver libro"
         Height          =   700
         Left            =   720
         TabIndex        =   17
         Top             =   8240
         Width           =   2600
      End
      Begin VB.CommandButton cmdAddFavorites 
         Caption         =   "Agregar a favoritos"
         Height          =   700
         Left            =   720
         TabIndex        =   16
         Top             =   7240
         Width           =   2600
      End
      Begin VB.Frame Frame2 
         Caption         =   "AddBooks"
         Height          =   4575
         Left            =   4080
         TabIndex        =   18
         Top             =   1200
         Width           =   12015
         Begin VB.CommandButton cmdClear 
            Caption         =   "Limpiar Campos"
            Height          =   700
            Left            =   7440
            TabIndex        =   13
            Top             =   240
            Width           =   2600
         End
         Begin VB.CommandButton cmdAddBook 
            Caption         =   "Agregar Libro"
            Height          =   700
            Left            =   3480
            TabIndex        =   12
            Top             =   240
            Width           =   2600
         End
         Begin VB.TextBox textDescription 
            Height          =   2490
            Left            =   7200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Text            =   "FormMain.frx":0000
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox textGenre 
            Height          =   450
            Left            =   1680
            TabIndex        =   10
            Text            =   "Genre"
            Top             =   3165
            Width           =   3500
         End
         Begin VB.TextBox textTitle 
            Height          =   450
            Left            =   1680
            TabIndex        =   7
            Text            =   "Title"
            Top             =   1365
            Width           =   3500
         End
         Begin VB.TextBox textAuthor 
            Height          =   450
            Left            =   1680
            TabIndex        =   8
            Text            =   "Author"
            Top             =   1965
            Width           =   3500
         End
         Begin VB.TextBox textYear 
            Height          =   450
            Left            =   1680
            TabIndex        =   9
            Text            =   "Year"
            Top             =   2565
            Width           =   3500
         End
         Begin VB.Label Label3 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Lucida Fax"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5520
            TabIndex        =   36
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label Label2 
            Caption         =   "Género"
            BeginProperty Font 
               Name            =   "Lucida Fax"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   35
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label lblTitle 
            Caption         =   "Título"
            BeginProperty Font 
               Name            =   "Lucida Fax"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   22
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label lblAuthor 
            Caption         =   "Autor"
            BeginProperty Font 
               Name            =   "Lucida Fax"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   21
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label lblyear 
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "Lucida Fax"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   20
            Top             =   2640
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Buscar libro"
         Height          =   700
         Left            =   720
         TabIndex        =   15
         Top             =   6240
         Width           =   2600
      End
      Begin VB.CommandButton cmdDeleteBook 
         Caption         =   "Eliminar Libro"
         Height          =   700
         Left            =   720
         TabIndex        =   19
         Top             =   9240
         Width           =   2600
      End
      Begin MSComctlLib.ListView ListBooks 
         Height          =   3975
         Left            =   4080
         TabIndex        =   23
         Top             =   6240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image imgCoverImg 
         Height          =   3000
         Left            =   600
         Top             =   2040
         Width           =   3000
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    initializeTextBox
End Sub

Private Sub cmdDeleteCompleted_Click()
    If Not ListCompletedBooks.SelectedItem Is Nothing Then
        
        Dim completedId As Integer
        completedId = ListCompletedBooks.SelectedItem.SubItems(6)
        
        Debug.Print "book id: " & completedId
        
        Dim sqlQuery As String
        sqlQuery = "delete from completed where id_completed =" & completedId
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        ListCompletedBooks.ListItems.Remove ListCompletedBooks.SelectedItem.Index
        MsgBox "Eliminado de libros Completados", vbInformation
        
        Exit Sub
        
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
End Sub


Private Sub cmdDeleteReading_Click()
    If Not ListReadings.SelectedItem Is Nothing Then
        
        Dim readingId As Integer
        readingId = ListReadings.SelectedItem.SubItems(6)
        
        'Debug.Print "book id: " & readingId
        
        Dim sqlQuery As String
        sqlQuery = "delete from readings where id_reading =" & readingId
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        ListReadings.ListItems.Remove ListReadings.SelectedItem.Index
        MsgBox "Eliminado de lista de libro leyendo", vbInformation
        
        Exit Sub
        
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
End Sub

Private Sub cmdDeleteNoWished_Click()
        If Not ListNoWished.SelectedItem Is Nothing Then
        
        Dim nowishedId As Integer
        nowishedId = ListNoWished.SelectedItem.SubItems(6)
        
        'Debug.Print "book id: " & readingId
        
        Dim sqlQuery As String
        sqlQuery = "delete from nowished where id_nowished =" & nowishedId
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        ListNoWished.ListItems.Remove ListNoWished.SelectedItem.Index
        MsgBox "Eliminado de lista de libros no deseados", vbInformation
        
        Exit Sub
        
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
End Sub

Private Sub cmdPerfil_Click()
    FormUserProfile.Show
    FormMain.Hide
    
End Sub

Private Sub cmdBookView_Click()
    
    If Not ListBooks.SelectedItem Is Nothing Then
        
        Set bookProperties = New book
        
        Dim id_book As Integer
        id_book = ListBooks.SelectedItem.SubItems(5)
        
        Dim sqlQuery As String
        sqlQuery = queryGetSpecificBook(id_book)
        
        Debug.Print sqlQuery
        
        Set rs = New ADODB.Recordset 'Activamos el Recordset
        'Abrimos la base de datos
        cn.Open connectionData
        rs.Source = "books" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
        rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
        rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
        rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
        Debug.Print rs.Fields("title")
    
        'Dim itm As ListItem
    
        'rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
        'Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    
        ' Ingresar datos en el objeto publico bookProperties
        bookProperties.title = rs.Fields("title")
        bookProperties.author = rs.Fields("author_name")
        bookProperties.year = rs.Fields("year")
        bookProperties.genre = rs.Fields("genres")
        bookProperties.description = rs.Fields("description")
        bookProperties.id_book = rs.Fields("id_book")
        
        Debug.Print bookProperties.title
        
        'rs.MoveNext 'Nos movemos al siguiente registro.
        'Loop

        ' Cerrar el recordset y la conexión
        rs.Close
        cn.Close

        ' Limpiar objetos
        Set rs = Nothing
        Set cn = Nothing
        
        viewBook
        
        FormBook.Show
        FormMain.Hide
        
    Else
        MsgBox "Selecciona un libro para visualizar"
    End If
    
    
End Sub

Private Sub cmdReadings_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, False, True, False
    
    Dim sqlQuery As String
    Debug.Print user.Id
    sqlQuery = queryGetReadings(user.Id)
    
    'sqlQuery = queryGetFavorites(1)
    Set rs = New ADODB.Recordset 'Activamos el Recordset

    'Abrimos la base de datos
    cn.Open connectionData
    rs.Source = "readings" 'Especificamos la fuente de datos. En este caso la tabla.
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    
    If rs.BOF And rs.EOF Then
    MsgBox "No hay libros agregados en lista de leyendo."
    
    Else
    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Debug.Print rs.Fields("id_book")
    
    Dim itm As ListItem
    ListReadings.ListItems.Clear
    
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListReadings.ListItems.Add(, , rs.Fields("title"))
    itm.SubItems(1) = rs.Fields("author_name")
    itm.SubItems(2) = rs.Fields("year")
    itm.SubItems(3) = rs.Fields("genre")
    itm.SubItems(4) = rs.Fields("description")
    itm.SubItems(5) = rs.Fields("id_book")
    itm.SubItems(6) = rs.Fields("id_reading")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop
    
    End If

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    
    ' Inicializar la colección
    'Set favorites = New Collection
    
    ' Función que ajusta tamaño y posición de controles
    PositionFrames
    
    ' Funcion para inicializar ListViews
    initializeListViews
    
    ' Funcion para inicializar TextBox
    initializeTextBox
    
    ' Funcion para mostrar u ocultar los frames que son llamados
    'hideShowFrames True, False, False, False, False
    
    ' Cargar libros de la Db a Listview 'ListBooks'
    cmdHome_Click
    
End Sub

Private Sub Form_Resize()
     ' Ajustar la posición y tamaño del Frame cuando se cambia el tamaño del formulario
    On Error Resume Next ' Manejar posibles errores al redimensionar
    
    If Me.WindowState <> vbMinimized Then
        fraNavbar.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Fra_NavBar
    
        fraHome.Width = Me.ScaleWidth - 3090 ' Ajustar el ancho del Frame
        fraHome.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Frame
        
        fraFavorites.Width = Me.ScaleWidth - 3090 ' Ajustar el ancho del Frame
        fraFavorites.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Frame
        
        fraCompletedBooks.Width = Me.ScaleWidth - 3090 ' Ajustar el ancho del Frame
        fraCompletedBooks.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Frame
        
        fraHistory.Width = Me.ScaleWidth - 3090 ' Ajustar el ancho del Frame
        fraHistory.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Frame
        
        frmNoWished.Width = Me.ScaleWidth - 3090 ' Ajustar el ancho del Frame
        frmNoWished.Height = Me.ScaleHeight - 50 ' Ajustar la altura del Frame
    End If
End Sub

Private Sub cmdAddFavorites_Click()
    
    If Not ListBooks.SelectedItem Is Nothing Then
        
        Dim favoriteItem As book
        Set favoriteItem = New book
        
        favoriteItem.id_book = ListBooks.SelectedItem.SubItems(5)
        
        Dim sqlQuery As String
        sqlQuery = queryAddFavorite(user.Id, favoriteItem.id_book)
    
        Set cmd = New ADODB.Command 'Activamos el command
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        MsgBox "Libro añadido a favoritos", vbInformation
        
    Else
        MsgBox "Selecciona un libro para añadir a favoritos"
    End If
    
End Sub

Private Sub ShowBooks()
    Dim i As Integer
    Dim book As book
    
    For i = 1 To favorites.Count
        Set book = favorites(i)
        Debug.Print "Title: " & book.title & ", Author: " & book.author; ", Genre: " & book.genre
    Next i
End Sub



Private Sub cmdHome_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames True, False, False, False, False
    
    Dim sqlQuery As String
    sqlQuery = queryGetBooks

    Set rs = New ADODB.Recordset 'Activamos el Recordset

    'Abrimos la base de datos
    cn.Open connectionData
    rs.Source = "books" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    Debug.Print rs.Fields("id_book")
    
    Dim itm As ListItem
    ListBooks.ListItems.Clear
    
    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListBooks.ListItems.Add(, , rs.Fields("title"))
    itm.SubItems(1) = rs.Fields("author_name")
    itm.SubItems(2) = rs.Fields("year")
    itm.SubItems(3) = rs.Fields("genres")
    itm.SubItems(4) = rs.Fields("description")
    itm.SubItems(5) = rs.Fields("id_book")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub cmdCompletedBooks_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, True, False, False
    
    Debug.Print user.Id
    Dim sqlQuery As String
    sqlQuery = queryGetCompleted(user.Id)

    'Set rs = New ADODB.Recordset 'Activamos el Recordset

    'Abrimos la base de datos
    cn.Open connectionData
    rs.Source = "completed" 'Especificamos la fuente de datos. En este caso la tabla.
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    If rs.BOF And rs.EOF Then
    MsgBox "No hay libros completados aún."
    
    Else
    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Debug.Print "id completed book: " & rs.Fields("id_completed")
    
    Dim itm As ListItem
    ListCompletedBooks.ListItems.Clear
    
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListCompletedBooks.ListItems.Add(, , rs.Fields("title"))
    itm.SubItems(1) = rs.Fields("author_name")
    itm.SubItems(2) = rs.Fields("year")
    itm.SubItems(3) = rs.Fields("genre")
    itm.SubItems(4) = rs.Fields("description")
    itm.SubItems(5) = rs.Fields("id_book")
    itm.SubItems(6) = rs.Fields("id_completed")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop
    
    End If

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
    
End Sub

Private Sub cmdHistory_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, False, True, False
End Sub

Private Sub cmdNoWished_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, False, False, True
    
    Dim sqlQuery As String
    Debug.Print user.Id
    sqlQuery = queryGetNoWished(user.Id)
    
    Set rs = New ADODB.Recordset 'Activamos el Recordset

    'Abrimos la base de datos
    cn.Open connectionData
    rs.Source = "favorites" 'Especificamos la fuente de datos. En este caso la tabla.
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    If rs.BOF And rs.EOF Then
    MsgBox "No hay libros agregados a no deseados aún."
    
    Else
    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Debug.Print rs.Fields("id_book")
    
    Dim itm As ListItem
    ListNoWished.ListItems.Clear
    
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListNoWished.ListItems.Add(, , rs.Fields("title"))
    itm.SubItems(1) = rs.Fields("author_name")
    itm.SubItems(2) = rs.Fields("year")
    itm.SubItems(3) = rs.Fields("genre")
    itm.SubItems(4) = rs.Fields("description")
    itm.SubItems(5) = rs.Fields("id_book")
    itm.SubItems(6) = rs.Fields("id_nowished")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop
    
    End If

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
    
End Sub

Private Sub cmdAddBook_Click()
    Dim newBook As New book
    newBook.title = textTitle.Text
    newBook.author = textAuthor.Text
    newBook.year = textYear.Text
    newBook.genre = textGenre.Text
    newBook.description = textDescription.Text
    newBook.CoverImg = imgCoverImg.Picture
    
    Dim sqlQuery As String
    sqlQuery = queryAddBook(newBook.title, newBook.year, newBook.description, newBook.author, newBook.genre)
    
    Set cmd = New ADODB.Command 'Activamos el command
    
    'Abrimos la base de datos
    cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
    
    cmd.ActiveConnection = cn
    cmd.CommandText = sqlQuery
    cmd.Execute

    ' Cerrar la conexión
    cn.Close
    
    MsgBox "Libro " & newBook.title & " añadido", , "Libro agregado"
    
    Dim itm As ListItem
    Set itm = ListBooks.ListItems.Add(, , newBook.title)
    itm.SubItems(1) = newBook.author
    itm.SubItems(2) = newBook.year
    itm.SubItems(3) = newBook.genre
    itm.SubItems(4) = newBook.description
    
    initializeTextBox

End Sub

Private Sub cmdFavorites_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, True, False, False, False
    
    Dim sqlQuery As String
    Debug.Print user.Id
    sqlQuery = queryGetFavorites(user.Id)
    
    'sqlQuery = queryGetFavorites(1)
    Set rs = New ADODB.Recordset 'Activamos el Recordset

    'Abrimos la base de datos
    cn.Open connectionData
    rs.Source = "favorites" 'Especificamos la fuente de datos. En este caso la tabla.
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open sqlQuery, cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    If rs.BOF And rs.EOF Then
    MsgBox "No hay favoritos agregados aún."
    
    Else
    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Debug.Print rs.Fields("id_book")
    
    Dim itm As ListItem
    ListFavorites.ListItems.Clear
    
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListFavorites.ListItems.Add(, , rs.Fields("title"))
    itm.SubItems(1) = rs.Fields("author_name")
    itm.SubItems(2) = rs.Fields("year")
    itm.SubItems(3) = rs.Fields("genre")
    itm.SubItems(4) = rs.Fields("description")
    itm.SubItems(5) = rs.Fields("id_book")
    itm.SubItems(6) = rs.Fields("id_favorite")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop
    
    End If

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Private Sub cmdDeleteBook_Click()
    If Not ListBooks.SelectedItem Is Nothing Then
        
        Debug.Print ListBooks.SelectedItem.Text
        
        Dim booktitle As String
        booktitle = ListBooks.SelectedItem.Text
        
        Dim sqlQuery As String
        sqlQuery = queryDeleteBooks(booktitle)
        
        Debug.Print sqlQuery
        
        'Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
    
        MsgBox "Libro " & booktitle & " eliminado", , "Libros"
        
        ListBooks.ListItems.Remove ListBooks.SelectedItem.Index
        
        Exit Sub
        
    Else
        MsgBox "Selecciona un libro para eliminar", vbInformation
    End If
    
End Sub

Private Sub cmdDeleteFavorite_Click()
    If Not ListFavorites.SelectedItem Is Nothing Then
        
        Dim bookId As Integer
        bookId = ListFavorites.SelectedItem.SubItems(6)
        
        Debug.Print "book id" & bookId
        
        Dim sqlQuery As String
        sqlQuery = "delete from favorites where id_favorite =" & bookId
        
        'Set cmd = New ADODB.Command 'Activamos el command
    
        'Abrimos la base de datos
        cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
        cmd.ActiveConnection = cn
        cmd.CommandText = sqlQuery
        cmd.Execute

        ' Cerrar la conexión
        cn.Close
        
        ListFavorites.ListItems.Remove ListFavorites.SelectedItem.Index
        MsgBox "Libro eliminado de favoritos", vbInformation
        
        Exit Sub
        
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
End Sub


Private Sub cmdSearch_Click()
    Dim searchTitle As String
    Dim item As ListItem
    
    searchTitle = InputBox("Introduce el titulo del libro a buscar", "Buscar Libro")
    
    Debug.Print searchTitle
    
    If searchTitle = Empty Then
       MsgBox "Ingresa el título del libro a buscar"
    Else
    For Each item In ListBooks.ListItems
        If LCase(item.Text) = LCase(searchTitle) Then
            MsgBox "Libro encontrado"
            ' Selecciona el elemento
            ListBooks.ListItems(item.Index).Selected = True
            ' Asegura que el elemento esté visible
            ListBooks.ListItems(item.Index).EnsureVisible
            ' Establece el foco en el ListView
            ListBooks.SetFocus
            
            Exit Sub
        End If
        Next
            MsgBox "Libro no encontrado"
    End If
End Sub


