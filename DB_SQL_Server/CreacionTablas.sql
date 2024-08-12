use HubDeLectura

-- Cración de entidad users
create table users (
	id_user int not null identity(1,1),
	name varchar(50) not null,
	last_name varchar(50),
	email varchar(30) not null,
	password varchar(30) not null,
	age smallint,
	photo varchar(1000),

	-- Restricciones
	constraint PK_id_user primary key (
		id_user asc
	),
	constraint Uq_email unique(
		email
	),
	constraint No_negative_age check(
		age > 0
	)
)

-- creacion de entidad authors
create table authors (
	id_author int not null identity(1,1),
	name varchar(50) not null,
	last_name varchar(50),
	nationality varchar(50),

	-- Restricciones
	constraint PK_id_author primary key (
		id_author asc
	)
)

-- Creación de entidad books
create table books (
	id_book int not null identity(1,1),
	title varchar(50) not null,
	year varchar(20),
	description varchar(5000),
	cover_img varchar(1000),
	id_author int,

	-- Restricciones
	constraint PK_id_book primary key (
		id_book asc
	),
	constraint FK_id_author foreign key (id_author) references authors(id_author),
	constraint Uq_book unique(title, year)
)

-- creación de entidad genres
create table genres (
	id_genre int not null identity(1,1),
	genre varchar(50),

	-- Restricciones
	constraint PK_id_genre primary key (
		id_genre asc
	) 
	--constraint Uq_genre unique(genre)
)

-- cración de entidad book_genres
create table book_genres(
	id_book_genre int not null identity(1,1),
	id_book int not null,
	id_genre int not null,

	-- Restricciones
	constraint PK_id_book_genre primary key (
		id_book_genre asc
	),
	constraint FK_id_bookgenre foreign key (id_book) references books(id_book),
	constraint FK_id_genrebook foreign key (id_genre) references genres(id_genre),
	--constraint Uq_bookgenres unique(id_genre, id_book)
	
)

-- creación de entidad favorites
create table favorites(
	id_favorite int not null identity(1,1),
	id_user int not null,
	id_book int not null,

	-- Restricciones
	constraint PK_id_favorite primary key (
		id_favorite asc
	),
	constraint FK_id_userfavorites foreign key (id_user) references users(id_user),
	constraint FK_id_bookfavorites foreign key (id_book) references books(id_book),
	--constraint Uq_bookfavorites unique(id_user, id_book)
)

-- creación de entidad Readings
create table readings(
	id_reading int not null identity(1,1),
	id_user int not null,
	id_book int not null,

	-- Restricciones
	constraint PK_id_reading primary key (
		id_reading asc
	),
	constraint FK_id_userreading foreign key (id_user) references users(id_user),
	constraint FK_id_bookreading foreign key (id_book) references books(id_book),
	constraint Uq_bookreading unique(id_user, id_book)
)

-- creación de entidad completedBooks
create table completed(
	id_completed int not null identity(1,1),
	id_user int not null,
	id_book int not null,

	-- Restricciones
	constraint PK_id_completed primary key (
		id_completed asc
	),
	constraint FK_id_usercompleted foreign key (id_user) references users(id_user),
	constraint FK_id_bookcompleted foreign key (id_book) references books(id_book),
	--constraint Uq_bookcompleted unique(id_user, id_book)
)

-- creación de entidad nowished
create table nowished(
	id_nowished int not null identity(1,1),
	id_user int not null,
	id_book int not null,

	-- Restricciones
	constraint PK_id_nowished primary key (
		id_nowished asc
	),
	constraint FK_id_usernowished foreign key (id_user) references users(id_user),
	constraint FK_id_booknowished foreign key (id_book) references books(id_book),
	--constraint Uq_booknowished unique(id_user, id_book)
)





