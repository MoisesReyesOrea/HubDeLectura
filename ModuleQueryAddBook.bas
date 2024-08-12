Attribute VB_Name = "ModuleQueriesSQL"

Function queryAddBook(title As String, year As String, description As String, authorName As String, genre As String) As String

queryAddBook = "BEGIN TRANSACTION;" & _
"INSERT INTO authors (name)" & _
"VALUES ('" & authorName & "');" & _
"DECLARE @AuthorID INT;" & _
"SET @AuthorID = SCOPE_IDENTITY();" & _
"INSERT INTO genres (genre)" & _
"VALUES ('" & genre & "');" & _
"DECLARE @GenreID INT;" & _
"SET @GenreID = SCOPE_IDENTITY();" & _
"INSERT INTO books (title, year, description, id_author)" & _
"VALUES ('" & title & "', '" & year & "', '" & description & "', @AuthorID );" & _
"DECLARE @BookID INT;" & _
"SET @BookID = SCOPE_IDENTITY();" & _
"INSERT INTO book_genres (id_book, id_genre)" & _
"VALUES (@BookID, @GenreID);" & _
"COMMIT TRANSACTION;"

End Function

Function queryAddFavorite(ByVal id_user As Integer, id_book As Integer) As String
    queryAddFavorite = "insert into favorites (id_book, id_user) " & _
    "values (" & id_book & ", " & id_user & ")"

End Function

Function queryAddCompleted(ByVal id_user As Integer, id_book As Integer) As String
    queryAddCompleted = "insert into completed (id_book, id_user) " & _
    "values (" & id_book & ", " & id_user & ")"

End Function

Function queryAddReading(ByVal id_user As Integer, id_book As Integer) As String
    queryAddReading = "insert into readings (id_book, id_user) " & _
    "values (" & id_book & ", " & id_user & ")"

End Function

Function queryAddNowished(ByVal id_user As Integer, id_book As Integer) As String
    queryAddNowished = "insert into nowished (id_book, id_user) " & _
    "values (" & id_book & ", " & id_user & ")"

End Function

Function queryGetBooks()

    queryGetBooks = "SELECT " & _
    "b.id_book, " & _
    "b.title, " & _
    "a.name AS author_name, " & _
    "b.year, " & _
    "STRING_AGG(g.genre, ', ') AS genres, " & _
    "b.description " & _
    "From " & _
    "books b " & _
    "Join " & _
    "authors a ON b.id_author = a.id_author " & _
    "Join " & _
    "book_genres bg ON b.id_book = bg.id_book " & _
    "Join " & _
    "genres g ON bg.id_genre = g.id_genre " & _
    "Group By " & _
    "b.id_book, b.title, b.year, b.description, " & _
    "a.Name " & _
    "Order By " & _
    "b.id_book; "

End Function

Function queryGetSpecificBook(ByVal id_book As Integer) As String
    
    queryGetSpecificBook = "SELECT " & _
    "b.id_book, " & _
    "b.title, " & _
    "a.name AS author_name, " & _
    "b.year, " & _
    "STRING_AGG(g.genre, ', ') AS genres, " & _
    "b.Description " & _
    "From " & _
    "books b " & _
    "Join " & _
    "authors a ON b.id_author = a.id_author " & _
    "Join " & _
    "book_genres bg ON b.id_book = bg.id_book " & _
    "Join " & _
    "genres g ON bg.id_genre = g.id_genre " & _
    "Where b.id_book =" & id_book & _
    " Group By " & _
    "b.id_book, b.title, b.year, b.description, " & _
    "a.Name"
    
End Function


Function queryGetFavorites(ByVal id_user As Integer) As String
    queryGetFavorites = "SELECT " & _
    "f.id_favorite, " & _
    "b.id_book, " & _
    "b.title, " & _
    "b.year, " & _
    "a.name AS author_name, " & _
    "g.genre, " & _
    "b.description " & _
    "FROM favorites f " & _
    "JOIN books b ON f.id_book = b.id_book " & _
    "JOIN authors a ON b.id_author = a.id_author " & _
    "JOIN book_genres bg ON b.id_book = bg.id_book " & _
    "JOIN genres g ON bg.id_genre = g.id_genre " & _
    "Where f.id_user = " & id_user & _
    "ORDER BY b.id_book, g.genre; "

End Function

Function queryGetCompleted(ByVal id_user As Integer) As String
    queryGetCompleted = "SELECT " & _
    "c.id_completed, " & _
    "b.id_book, " & _
    "b.title, " & _
    "b.year, " & _
    "a.name AS author_name, " & _
    "g.genre, " & _
    "b.description " & _
    "FROM completed c " & _
    "JOIN books b ON c.id_book = b.id_book " & _
    "JOIN authors a ON b.id_author = a.id_author " & _
    "JOIN book_genres bg ON b.id_book = bg.id_book " & _
    "JOIN genres g ON bg.id_genre = g.id_genre " & _
    "Where c.id_user = " & id_user & _
    "ORDER BY b.id_book, g.genre; "
    
End Function

Function queryGetReadings(ByVal id_user As Integer) As String
    queryGetReadings = "SELECT " & _
    "r.id_reading, " & _
    "b.id_book, " & _
    "b.title, " & _
    "b.year, " & _
    "a.name AS author_name, " & _
    "g.genre, " & _
    "b.description " & _
    "FROM readings r " & _
    "JOIN books b ON r.id_book = b.id_book " & _
    "JOIN authors a ON b.id_author = a.id_author " & _
    "JOIN book_genres bg ON b.id_book = bg.id_book " & _
    "JOIN genres g ON bg.id_genre = g.id_genre " & _
    "Where r.id_user = " & id_user & _
    "ORDER BY b.id_book, g.genre; "
End Function

Function queryGetNoWished(ByVal id_user As Integer) As String
    queryGetNoWished = "SELECT " & _
    "nw.id_nowished, " & _
    "b.id_book, " & _
    "b.title, " & _
    "b.year, " & _
    "a.name AS author_name, " & _
    "g.genre, " & _
    "b.description " & _
    "FROM nowished nw " & _
    "JOIN books b ON nw.id_book = b.id_book " & _
    "JOIN authors a ON b.id_author = a.id_author " & _
    "JOIN book_genres bg ON b.id_book = bg.id_book " & _
    "JOIN genres g ON bg.id_genre = g.id_genre " & _
    "Where nw.id_user = " & id_user & _
    "ORDER BY b.id_book, g.genre; "
End Function


Function queryDeleteBooks(ByVal booktitle As String) As String

queryDeleteBooks = "BEGIN TRANSACTION;" & _
"DECLARE @title VARCHAR(50);" & _
"SET @title = '" & booktitle & "' ;" & _
"DELETE FROM book_genres" & _
" WHERE id_book IN (SELECT id_book FROM books WHERE title = @title);" & _
"DELETE FROM favorites" & _
" WHERE id_book IN (SELECT id_book FROM books WHERE title = @title);" & _
"DELETE FROM readings" & _
" WHERE id_book IN (SELECT id_book FROM books WHERE title = @title);" & _
"DELETE FROM completed" & _
" WHERE id_book IN (SELECT id_book FROM books WHERE title = @title);" & _
"DELETE FROM nowished" & _
" WHERE id_book IN (SELECT id_book FROM books WHERE title = @title);" & _
"DELETE FROM books" & _
" WHERE title = @title;" & _
"COMMIT TRANSACTION;"

End Function


