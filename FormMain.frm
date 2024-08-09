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
   Begin VB.Frame fraFavorites 
      Caption         =   "Favorites"
      Height          =   3735
      Left            =   3000
      TabIndex        =   16
      Top             =   7200
      Width           =   16455
      Begin VB.CommandButton cmdDeleteFavorite 
         Caption         =   "Eliminar favorito"
         Height          =   855
         Left            =   840
         TabIndex        =   29
         Top             =   360
         Width           =   2895
      End
      Begin MSComctlLib.ListView ListFavorites 
         Height          =   2295
         Left            =   4560
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.Frame fraHistory 
      Caption         =   "History"
      Height          =   4935
      Left            =   14400
      TabIndex        =   35
      Top             =   6960
      Width           =   15495
   End
   Begin VB.Frame frmNoWished 
      Caption         =   "NoWished"
      Height          =   5055
      Left            =   20040
      TabIndex        =   32
      Top             =   2640
      Width           =   15255
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   855
         Left            =   360
         TabIndex        =   34
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   2760
         TabIndex        =   33
         Top             =   1800
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraCompletedBooks 
      Caption         =   "CompletedBooks"
      Height          =   5175
      Left            =   4080
      TabIndex        =   27
      Top             =   5760
      Width           =   16815
      Begin VB.CommandButton cmdDatosDB 
         Caption         =   "DatosDB"
         Height          =   1095
         Left            =   360
         TabIndex        =   31
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CommandButton cmdDeleteCompleted 
         Caption         =   "Eliminar de Completados"
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
         TabIndex        =   25
         Top             =   9800
         Width           =   2175
      End
      Begin VB.CommandButton cmdGenres 
         Caption         =   "Generos"
         Height          =   1000
         Left            =   360
         TabIndex        =   24
         Top             =   8300
         Width           =   2175
      End
      Begin VB.CommandButton cmdNoWished 
         Caption         =   "No deseados"
         Height          =   1000
         Left            =   360
         TabIndex        =   23
         Top             =   6800
         Width           =   2175
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "Leyendo"
         Height          =   1000
         Left            =   360
         TabIndex        =   22
         Top             =   5300
         Width           =   2175
      End
      Begin VB.CommandButton cmdFavorites 
         Appearance      =   0  'Flat
         Caption         =   "Favoritos"
         Height          =   1000
         Left            =   360
         TabIndex        =   21
         Top             =   2300
         Width           =   2175
      End
      Begin VB.CommandButton cmdHome 
         Caption         =   "Inicio"
         Height          =   1000
         Left            =   360
         TabIndex        =   20
         Top             =   800
         Width           =   2175
      End
      Begin VB.CommandButton cmdCompletedBooks 
         Caption         =   "Leidos"
         Height          =   1000
         Left            =   360
         TabIndex        =   19
         Top             =   3800
         Width           =   2175
      End
   End
   Begin VB.Frame fraHome 
      Caption         =   "Home"
      Height          =   10635
      Left            =   3060
      TabIndex        =   8
      Top             =   0
      Width           =   16575
      Begin VB.CommandButton Command2 
         Caption         =   "Ver libro"
         Height          =   850
         Left            =   720
         TabIndex        =   38
         Top             =   8240
         Width           =   2800
      End
      Begin VB.CommandButton cmdAddFavorites 
         Caption         =   "Agregar a favoritos"
         Height          =   850
         Left            =   720
         TabIndex        =   26
         Top             =   7240
         Width           =   2800
      End
      Begin VB.Frame Frame2 
         Caption         =   "AddBooks"
         Height          =   4575
         Left            =   4080
         TabIndex        =   11
         Top             =   1200
         Width           =   12015
         Begin VB.CommandButton cmdClear 
            Caption         =   "Limpiar Campos"
            Height          =   850
            Left            =   7440
            TabIndex        =   7
            Top             =   240
            Width           =   2800
         End
         Begin VB.CommandButton cmdAddBook 
            Caption         =   "Agregar Libro"
            Height          =   850
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   2800
         End
         Begin VB.TextBox textDescription 
            Height          =   2490
            Left            =   7200
            TabIndex        =   5
            Text            =   "Description"
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox textGenre 
            Height          =   450
            Left            =   1680
            TabIndex        =   4
            Text            =   "Genre"
            Top             =   3165
            Width           =   3500
         End
         Begin VB.TextBox textTitle 
            Height          =   450
            Left            =   1680
            TabIndex        =   1
            Text            =   "Title"
            Top             =   1365
            Width           =   3500
         End
         Begin VB.TextBox textAuthor 
            Height          =   450
            Left            =   1680
            TabIndex        =   2
            Text            =   "Author"
            Top             =   1965
            Width           =   3500
         End
         Begin VB.TextBox textYear 
            Height          =   450
            Left            =   1680
            TabIndex        =   3
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   2640
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Buscar libro"
         Height          =   850
         Left            =   720
         TabIndex        =   10
         Top             =   6240
         Width           =   2800
      End
      Begin VB.CommandButton cmdDeleteBook 
         Caption         =   "Eliminar Libro"
         Height          =   850
         Left            =   720
         TabIndex        =   9
         Top             =   9240
         Width           =   2800
      End
      Begin MSComctlLib.ListView ListBooks 
         Height          =   3975
         Left            =   4080
         TabIndex        =   15
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
'Dim cn As New ADODB.Connection  'Creamos el objeto Connection.
'Public WithEvents rs As ADODB.Recordset 'Creamos el Recordset con soporte de eventos.
'Dim rs As New ADODB.Recordset 'Creamos el objeto Recordset.

Option Explicit

Private Sub cmdPerfil_Click()
    FormUserProfile.Show
    FormMain.Hide
    
End Sub

'Private favorites As Collection

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
    hideShowFrames True, False, False, False, False
    
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
        
        Dim item As book
        Set item = New book
        item.Title = ListBooks.SelectedItem.Text
        item.Author = ListBooks.SelectedItem.SubItems(1)
        item.Genre = ListBooks.SelectedItem.SubItems(2)
        'item.CoverImg = ListBooks.SelectedItem.Picture
        
        'favorites.Add item
        MsgBox "Libro añadido a favoritos", vbInformation
        'ShowBooks
        
        Dim itm As ListItem
        Set itm = ListFavorites.ListItems.Add(, , item.Title)
        itm.SubItems(1) = item.Author
        itm.SubItems(2) = item.Genre
    Else
        MsgBox "Selecciona un libro para añadir a favoritos"
    End If
End Sub

Private Sub ShowBooks()
    Dim i As Integer
    Dim book As book
    
    For i = 1 To favorites.Count
        Set book = favorites(i)
        Debug.Print "Title: " & book.Title & ", Author: " & book.Author; ", Genre: " & book.Genre
    Next i
End Sub



Private Sub cmdHome_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames True, False, False, False, False
    
End Sub



Private Sub cmdCompletedBooks_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, True, False, False
    
End Sub

Private Sub cmdHistory_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, False, True, False
End Sub

Private Sub cmdNoWished_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, False, False, False, True
End Sub

Private Sub cmdAddBook_Click()
    Dim newBook As New book
    newBook.Title = textTitle.Text
    newBook.Author = textAuthor.Text
    newBook.Year = textYear.Text
    newBook.Genre = textGenre.Text
    newBook.Description = textDescription.Text
    newBook.CoverImg = imgCoverImg.Picture
    
    Dim itm As ListItem
    Set itm = ListBooks.ListItems.Add(, , newBook.Title)
    itm.SubItems(1) = newBook.Author
    itm.SubItems(2) = newBook.Year
    itm.SubItems(3) = newBook.Genre
    itm.SubItems(4) = newBook.Description
    
    initializeTextBox
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command 'Activamos el command
    
    'Dim connectionData As String
    'connectionData = "Provider=" + providerDB + ";Data Source=" + sourceDB + ";Initial Catalog=" + nameDB + ";User ID=" + userIdDB + ";Password=" + passDB + ";"
    'connectionData = "Provider=SQLOLEDB;Data Source=LAPTOPS1;Initial Catalog=HubDeLectura;User ID=usersql;Password=root;"

    'Abrimos la base de datos
    'cn.Open "Provider=" + SQLOLEDB + ";Data Source=" + LAPTOPS1 + ";Initial Catalog=" + HubDeLectura + ";User ID=" + usersql + ";Password=" + Root + ";"
    cn.Open connectionData ' Constante "connectionData" contiene las variables de entorno para la conexion a la DB archivo no mostrado
    
    cmd.ActiveConnection = cn
    cmd.CommandText = "INSERT INTO books (title, year, description, id_author) VALUES (?, ?, ?, ?)"

    cmd.Parameters.Append cmd.CreateParameter("title", adVarChar, adParamInput, 50, newBook.Title)
    cmd.Parameters.Append cmd.CreateParameter("year", adVarChar, adParamInput, 50, newBook.Year)
    cmd.Parameters.Append cmd.CreateParameter("description", adVarChar, adParamInput, 10000, newBook.Description)
    cmd.Parameters.Append cmd.CreateParameter("id_author", adVarChar, adParamInput, 10, 1)
    cmd.Execute

    ' Cerrar la conexión
    cn.Close
    
    MsgBox "Libro " & newBook.Title & " añadido", , "Libro agregado"
    
    Exit Sub

End Sub

Private Sub cmdDatosDB_Click()
Set rs = New ADODB.Recordset 'Activamos el Recordset

'Abrimos la base de datos
cn.Open connectionData
rs.Source = "books" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
rs.Open "select * from books", cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.

Dim itm As ListItem

rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
Set itm = ListCompletedBooks.ListItems.Add(, , rs.Fields("title"))
'itm.SubItems(1) = rs.Fields("author")
itm.SubItems(2) = rs.Fields("year")
'itm.SubItems(3) = rs.Fields("genre")
itm.SubItems(4) = rs.Fields("description")
rs.MoveNext 'Nos movemos al siguiente registro.
Loop

' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing

End Sub

Private Sub cmdFavorites_Click()
    ' Funcion para mostrar u ocultar los frames que son llamados
    hideShowFrames False, True, False, False, False
    
    Set rs = New ADODB.Recordset 'Activamos el Recordset
    'Abrimos la base de datos
    cn.Open "Provider=SQLOLEDB;Data Source=LAPTOPS1;Initial Catalog=HubDeLectura;User ID=usersql;Password=root;" '"Provider=SQLOLEDB;" & "Data Source=LAPTOPS1"
    rs.Source = "favorites" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open "select * from favorites f JOIN books b on f.id_book = b.id_book where id_user = '1'", cn 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    
    Debug.Print rs.Fields("title")
    
    Dim itm As ListItem

    rs.MoveFirst 'Nos posicionamos en el primer registro del Recordset.
    Do Until rs.EOF 'Repite hasta que se lea todo el Recordset.
    Set itm = ListFavorites.ListItems.Add(, , rs.Fields("title"))
    'itm.SubItems(1) = rs.Fields("author")
    itm.SubItems(2) = rs.Fields("year")
    'itm.SubItems(3) = rs.Fields("genre")
    itm.SubItems(4) = rs.Fields("description")
    rs.MoveNext 'Nos movemos al siguiente registro.
    Loop

    ' Cerrar el recordset y la conexión
    rs.Close
    cn.Close

    ' Limpiar objetos
    Set rs = Nothing
    Set cn = Nothing
    
    
End Sub

Private Sub cmdDeleteBook_Click()
    If Not ListBooks.SelectedItem Is Nothing Then
        ListBooks.ListItems.Remove ListBooks.SelectedItem.Index
        MsgBox "Libro eliminado", vbInformation
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
    
End Sub

Private Sub cmdDeleteFavorite_Click()
    If Not ListFavorites.SelectedItem Is Nothing Then
        ListFavorites.ListItems.Remove ListFavorites.SelectedItem.Index
        MsgBox "Libro eliminado de favoritos", vbInformation
    Else
        MsgBox "Selecciona un libro para eliminar"
    End If
End Sub


Private Sub cmdSearch_Click()
    Dim searchTitle As String
    Dim item As ListItem
    
    searchTitle = InputBox("Introduce el titulo del libro a buscar", "Buscar Libro")
    
    For Each item In ListBooks.ListItems
        If LCase(item.Text) = LCase(searchTitle) Then
            item.Selected = True
            item.EnsureVisible
            Exit For
        End If
        Next
        If ListBooks.SelectedItem Is Nothing Then
            MsgBox "Libro no encontrado"
        End If
End Sub


