use HubDeLectura

drop table history

ALTER TABLE favorites
DROP CONSTRAINT FK__favorites__id_us__47DBAE45;

ALTER TABLE genres
ALTER COLUMN genre VARCHAR(50);


ALTER TABLE favorites
add constraint FK_id_user
foreign key (id_user) references users(id_user);

-- Extraer información de los libros junto con el nombre de sus autores
SELECT 
    b.id_book,
    b.title,
    b.year,
    b.description,
    b.cover_img,
    a.name AS author_name,
    a.last_name AS author_last_name,
    a.nationality AS author_nationality
FROM 
    books b
JOIN 
    authors a
ON 
    b.id_author = a.id_author;



-- Extraer información de book, authors y generos
SELECT 
    b.id_book,
    b.title,
    b.year,
    a.name AS author_name,
    a.last_name AS author_last_name,
    a.nationality AS author_nationality,
    g.genre
FROM 
    books b
JOIN 
    authors a ON b.id_author = a.id_author
JOIN 
    book_genres bg ON b.id_book = bg.id_book
JOIN 
    genres g ON bg.id_genre = g.id_genre
ORDER BY 
    b.id_book, g.genre;



-- Extraer información de book, authors y generos en una sola fila
SELECT 
    b.id_book,
    b.title,
    b.year,
    b.description,
    b.cover_img,
    a.name AS author_name,
    a.last_name AS author_last_name,
    a.nationality AS author_nationality,
    STRING_AGG(g.genre, ', ') AS genres
FROM 
    books b
JOIN 
    authors a ON b.id_author = a.id_author
JOIN 
    book_genres bg ON b.id_book = bg.id_book
JOIN 
    genres g ON bg.id_genre = g.id_genre
GROUP BY 
    b.id_book, b.title, b.year, b.description, b.cover_img, 
    a.name, a.last_name, a.nationality
ORDER BY 
    b.id_book;

--SELECT: Especifica las columnas que deseas extraer.
--STRING_AGG(g.genre, ', ') AS genres: Usa la función STRING_AGG para concatenar los géneros en una sola cadena, separada por comas.
--JOIN: Combina filas de varias tablas basadas en condiciones específicas.
--GROUP BY: Agrupa los resultados por los campos del libro y del autor para asegurar que cada libro aparezca una sola vez con sus géneros concatenados.
--ORDER BY: Ordena los resultados por id_book.




-- Extraer información de book, authors y generos en una sola fila
SELECT 
    b.id_book,
    b.title,
	a.name AS author_name,
    b.year,
	STRING_AGG(g.genre, ', ') AS genres,
    b.description
FROM 
    books b
JOIN 
    authors a ON b.id_author = a.id_author
JOIN 
    book_genres bg ON b.id_book = bg.id_book
JOIN 
    genres g ON bg.id_genre = g.id_genre
GROUP BY 
    b.id_book, b.title, b.year, b.description,
    a.name
ORDER BY 
    b.id_book;



-- Extraer información de un especifico book, authors y generos en una sola fila
SELECT 
    b.id_book,
    b.title,
	a.name AS author_name,
    b.year,
	STRING_AGG(g.genre, ', ') AS genres,
    b.description
FROM 
    books b
JOIN 
    authors a ON b.id_author = a.id_author
JOIN 
    book_genres bg ON b.id_book = bg.id_book
JOIN 
    genres g ON bg.id_genre = g.id_genre
WHERE b.id_book = 11
GROUP BY 
    b.id_book, b.title, b.year, b.description,
    a.name




