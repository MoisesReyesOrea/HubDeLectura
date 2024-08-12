Attribute VB_Name = "ModuleFuntions"

Function hideShowFrames(fraHome As Boolean, fraFavorites As Boolean, fraCompletedBooks As Boolean, fraHistory As Boolean, frmNoWished As Boolean)
    FormMain.fraHome.Visible = fraHome
    FormMain.fraFavorites.Visible = fraFavorites
    FormMain.fraCompletedBooks.Visible = fraCompletedBooks
    FormMain.fraHistory.Visible = fraHistory
    FormMain.frmNoWished.Visible = frmNoWished
End Function

Function initializeListViews()
    With FormMain.ListBooks
        View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2200
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Generos", 2200
        .ColumnHeaders.Add , , "Descripción", 4200
        .ColumnHeaders.Add , , "Id", 1
        
    End With
    
    With FormMain.ListFavorites
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2200
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Generos", 2200
        .ColumnHeaders.Add , , "Descripción", 4200
        .ColumnHeaders.Add , , "Id", 1
        .ColumnHeaders.Add , , "IdFavorite", 1
        
    End With
    
    With FormMain.ListCompletedBooks
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2200
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Generos", 2200
        .ColumnHeaders.Add , , "Descripción", 4200
        .ColumnHeaders.Add , , "Id", 1
        .ColumnHeaders.Add , , "IdCompleted", 1
        
    End With
    
    With FormMain.ListReadings
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2200
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Generos", 2200
        .ColumnHeaders.Add , , "Descripción", 4200
        .ColumnHeaders.Add , , "Id", 1
        .ColumnHeaders.Add , , "IdBookReading", 1
        
    End With
    
    With FormMain.ListNoWished
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2500
        .ColumnHeaders.Add , , "Año", 800
        .ColumnHeaders.Add , , "Generos", 2200
        .ColumnHeaders.Add , , "Descripción", 4200
        .ColumnHeaders.Add , , "Id", 1
        .ColumnHeaders.Add , , "IdNoWished", 1
        
    End With
End Function


Function initializeTextBox()
    FormMain.textTitle.Text = ""
    FormMain.textAuthor.Text = ""
    FormMain.textYear.Text = ""
    FormMain.textGenre.Text = ""
    FormMain.textDescription.Text = ""
    FormMain.imgCoverImg.Picture = LoadPicture("")
End Function

Public Function viewBook()
    Debug.Print "en FormBook " & bookProperties.title
        
    'Me.lblTitleBook.Caption = ""
    
    Dim title As String
    title = "Título: " & bookProperties.title
    Dim author As String
    author = "Autor: " & bookProperties.author
    Dim year As String
    year = "Año: " & bookProperties.year
    Dim genre As String
    genre = "Género: " & bookProperties.genre
    Dim description As String
    description = "Descripción: " & bookProperties.description
    
    FormBook.lblTitleBook.Caption = title
    FormBook.lblAuthorBook.Caption = author
    FormBook.lblYearBook.Caption = year
    FormBook.lblGenreBook.Caption = genre
    FormBook.lblDescriptionBook.Caption = description
    
End Function









