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
	values ('Antoine', 'de Saint-Exup�ry', 'Francia')


-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('1984', '1949',
	'Un cl�sico sobre los totalitarismos y la manipulaci�n de la verdad de una vigencia escalofriante. Una de las novelas m�s atractivas e inquietantes del siglo XX.',
	1 )

-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('Rebeli�n en la granja', '1945',
	'Rebeli�n en la granja es una f�bula dist�pica del escritor ingl�s George Orwell. Orwell arma una cr�tica a Stalin a trav�s de la personificaci�n de los animales. Por lo tanto es un libro aleg�rico escrito y publicado durante la Segunda Guerra Mundial con un fuerte car�cter cuestionador.',
	1 )

-- Insertando datos en books
insert into books (title, year,	description, id_author)
	values ('El principito', '1943',
	'El autor se estrella con su avi�n en medio del desierto del Sahara y encuentra a un ni�o, que es un pr�ncipe de otro planeta. Se trata de un relato po�tico que es filos�fico e incluye cr�tica social. Hace diversas observaciones sobre la naturaleza humana y su lectura es placentera y al mismo tiempo invita a la reflexi�n. Y naturalmente, est� pensado para lectores de todas las edades. El principito es una narraci�n corta del escritor franc�s Antoine de Saint-Exup�ry. La historia se centra en un peque�o pr�ncipe que realiza una traves�a por el universo. En este viaje descubre la extra�a forma en que los adultos ven la vida y comprende el valor del amor y la amistad.',
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

-- A�adir restriccion a tabla books para evitar books repetidos
alter table books
add constraint Uq_book unique(title, year)




-- Insertar libro con su autor y su genero
BEGIN TRANSACTION;
-- Insertar el autor en la tabla authors
INSERT INTO authors (name)
VALUES ('R.R. Martin');
-- Obtener el id del autor reci�n insertado
DECLARE @AuthorID INT;
SET @AuthorID = SCOPE_IDENTITY();
-- Insertar el g�nero en la tabla genres
INSERT INTO genres (genre)
VALUES ('Alta fantas�a');
-- Obtener el id del g�nero reci�n insertado
DECLARE @GenreID INT;
SET @GenreID = SCOPE_IDENTITY();
-- Insertar el libro en la tabla books utilizando el ID del autor
INSERT INTO books (title, year, description, id_author)
VALUES (
    'Danza de dragones', 
    '2011', 
    'Danza de dragones es la quinta de la serie de siete novelas previstas en la serie de fantas�a �pica Canci�n de hielo y fuego del autor estadounidense George R. R. Martin.', 
    @AuthorID
);
-- Obtener el id del libro reci�n insertado
DECLARE @BookID INT;
SET @BookID = SCOPE_IDENTITY();
-- Insertar la relaci�n entre el libro y el g�nero en la tabla book_genres
INSERT INTO book_genres (id_book, id_genre)
VALUES (@BookID, @GenreID);
COMMIT TRANSACTION;


