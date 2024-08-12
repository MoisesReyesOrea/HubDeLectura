use HubDeLectura

select * from users
select * from authors
select * from books
select * from genres
select * from book_genres
select * from favorites
select * from readings
select * from completed
select * from nowished


select * from users where email = 'moises@gmail.com';


-- Comando para mostrar todas las tablas de la DB
SELECT TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE';

-- Comando para mostrar todas las tablas de la DB
SELECT name
FROM sys.tables;

-- Comando para mostrar todas las tablas de la DB con esquema al que pertenecen
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE';

-- Comando para mostrar todas las tablas de la DB con esquema al que pertenecen
SELECT s.name AS SchemaName, t.name AS TableName
FROM sys.tables t
INNER JOIN sys.schemas s ON t.schema_id = s.schema_id;


-- Extraer toda la informacion de favoritos
select * from favorites f
JOIN books b on f.id_book = b.id_book
where id_user = '1'


-- Extraer información selecta de tabla favoritos
Select
	b.id_book,
    b.title,
	b.year,
	f.id_favorite
from favorites f
Join books b
on f.id_book = b.id_book


-- Extraer información de favoritos con book, authors y generos
SELECT 
	f.id_favorite,
    b.id_book,
    b.title,
    b.year,
    a.name AS author_name,
    g.genre,
	b.description
FROM favorites f 
JOIN books b ON f.id_book = b.id_book
JOIN authors a ON b.id_author = a.id_author
JOIN book_genres bg ON b.id_book = bg.id_book
JOIN genres g ON bg.id_genre = g.id_genre
WHERE f.id_user = 3
ORDER BY b.id_book, g.genre;