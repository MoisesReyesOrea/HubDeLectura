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
        .ColumnHeaders.Add , , "Autor", 2500
        .ColumnHeaders.Add , , "Año", 1500
        .ColumnHeaders.Add , , "Generos", 2500
        .ColumnHeaders.Add , , "Descripción", 5000
        
    End With
    
    With FormMain.ListFavorites
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2500
        .ColumnHeaders.Add , , "Año", 1500
        .ColumnHeaders.Add , , "Generos", 2500
        .ColumnHeaders.Add , , "Descripción", 5000
        
    End With
    
    With FormMain.ListCompletedBooks
        .View = lvwReport
        .ColumnHeaders.Add , , "Titulo", 2500
        .ColumnHeaders.Add , , "Autor", 2500
        .ColumnHeaders.Add , , "Año", 1500
        .ColumnHeaders.Add , , "Generos", 2500
        .ColumnHeaders.Add , , "Descripción", 5000
        
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
