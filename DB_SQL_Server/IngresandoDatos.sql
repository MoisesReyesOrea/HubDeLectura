use HubDeLectura

-- Insertar datos de usuarios user test
insert into users (name, last_name, email, password, age)
	values ('Moises', 'Reyes', 'q', '1', 30)

-- Insertar datos de usuarios
insert into users (name, last_name, email, password, age)
	values ('Moises', 'Reyes', 'moises@gmail.com', '123456', 30)

-- Insertar datos de usuarios
insert into users (name, last_name, email, password, age)
	values ('Moises', 'Orea', 'orea@gmail.com', '123456', 28)


-- Insertando datos en Authors
insert into authors (name, last_name, nationality)
	values ('George', 'Orwell', 'India')

-- Insertando datos en Authors
insert into authors (name, last_name, nationality)
	values ('Antoine', 'de Saint-Exupéry', 'Francia')


-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('1984', '1949',
	'Un clásico sobre los totalitarismos y la manipulación de la verdad de una vigencia escalofriante. Una de las novelas más atractivas e inquietantes del siglo XX.',
	1 )

-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('Rebelión en la granja', '1945',
	'Rebelión en la granja es una fábula distópica del escritor inglés George Orwell. Orwell arma una crítica a Stalin a través de la personificación de los animales. Por lo tanto es un libro alegórico escrito y publicado durante la Segunda Guerra Mundial con un fuerte carácter cuestionador.',
	1 )

-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('El principito', '1943',
	'El autor se estrella con su avión en medio del desierto del Sahara y encuentra a un niño, que es un príncipe de otro planeta. Se trata de un relato poético que es filosófico e incluye crítica social. Hace diversas observaciones sobre la naturaleza humana y su lectura es placentera y al mismo tiempo invita a la reflexión. Y naturalmente, está pensado para lectores de todas las edades. El principito es una narración corta del escritor francés Antoine de Saint-Exupéry. La historia se centra en un pequeño príncipe que realiza una travesía por el universo. En este viaje descubre la extraña forma en que los adultos ven la vida y comprende el valor del amor y la amistad.',
	 2)


-- Insertando datos en Favorites
insert into favorites (id_book, id_user)
	values (11, 5)


-- Insertando datos en Genres
insert into genres (genre)
	values 
		('drama')
		--('ficcion')
		--('satira')


-- Insertando datos en book_Genres
insert into book_genres (id_book, id_genre)
	values (3,1)

-- Insertando datos en Readings
insert into readings (id_book, id_user)
	values (3, 3)

-- Insertando datos en completed
insert into completed (id_book, id_user)
	values (3, 3)

-- Insertando datos en nowished
insert into nowished (id_book, id_user)
	values (3, 3)

-- Añadir restriccion a tabla books para evitar books repetidos
alter table books
add constraint Uq_book unique(title, year)




-- Insertar libro con su autor y su genero
BEGIN TRANSACTION;
-- Insertar el autor en la tabla authors
INSERT INTO authors (name)
VALUES ('R.R. Martin');
-- Obtener el id del autor recién insertado
DECLARE @AuthorID INT;
SET @AuthorID = SCOPE_IDENTITY();
-- Insertar el género en la tabla genres
INSERT INTO genres (genre)
VALUES ('Alta fantasía');
-- Obtener el id del género recién insertado
DECLARE @GenreID INT;
SET @GenreID = SCOPE_IDENTITY();
-- Insertar el libro en la tabla books utilizando el ID del autor
INSERT INTO books (title, year, description, id_author)
VALUES (
    'Danza de dragones', 
    '2011', 
    'Danza de dragones es la quinta de la serie de siete novelas previstas en la serie de fantasía épica Canción de hielo y fuego del autor estadounidense George R. R. Martin.', 
    @AuthorID
);
-- Obtener el id del libro recién insertado
DECLARE @BookID INT;
SET @BookID = SCOPE_IDENTITY();
-- Insertar la relación entre el libro y el género en la tabla book_genres
INSERT INTO book_genres (id_book, id_genre)
VALUES (@BookID, @GenreID);
COMMIT TRANSACTION;


