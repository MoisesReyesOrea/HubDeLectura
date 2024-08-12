# Hub de lectura - Visual Basic 6.0, WinForm proyecto para Mega - Liderly

## 1. Descripción
Este repositorio contiene una aplicacion de escritorio para windows hecha con el lenguaje Visual Basic 6.0 y SQL Server (Transact SQL). Esta aplicación windows es acerca de un Hub de libros en la cual se pueden agregar usuarios, cada usuario puede seleccionar su contenido como Favoritos, libros ya leidos, libros que está leyendo, libros no deseados, ademas de poder visualizar es posible añadir, eliminar y modificar tanto usuarios como libros. 

### Objetivo
El objetivo es crear una aplicación windows con visual basic 6.0 en que el usuario pueda organizar y controlar los libros que desea leer.

## 2. Requerimientos técnicos:
### Para visualizar el contenido del codigo es necesario tener instalado:  
Microsoft Visual Basic 6.0 
GIT: Debe tener Instalado GIT  
SQL Server: En este caso la aplicación se conecta a una base de datos local realizada en SQL Server.  

## 3. ¿Cómo ejecutar la aplicación?

-- Clona el repositorio con el comando:  ```git clone https://github.com/MoisesReyesOrea/HubDeLectura.git```  
-- Desde Microsoft Visual basic 6.0 abre el proyecto en 'Open project' 
-- Este repositorio no contiene el archivo con las variables de entorno de las credenciales de la base de datos SQL Server, para eso debes crear un archivo 'module'
(ej: ModuleEnvirontmentVariable.bas) dentro del proyecto en la carpeta 'Modules' en ese archivo ingresa las credenciales de tu DB, datos necesarios:

Public Const providerDB As String = "SQLOLEDB"
Public Const sourceDB As String = "Nombre del servidor de tu base de datos"
Public Const nameDB As String = "Nombre de tu base de datos"
Public Const userIdDB As String = "tu usuario"
Public Const passDB As String = "tu contraseña"
Public Const connectionData As String = "Provider=" + providerDB + ";Data Source=" + sourceDB + ";Initial Catalog=" + nameDB + ";User ID=" + userIdDB + ";Password=" + passDB + ";"

-- Corre la aplicacion desde el mismo Microsoft Visual Basic 6.0 

### Ejecuta el archivo HubDeLecturaMega.exe
--   

## 4. Explicación

Para correr la aplicación se ejecuta el comando 'node index.js', esta imágen muestra el server en ejecucion y conectada correctamente a la base de datos.
![APIenEjecucion](https://github.com/user-attachments/assets/92cbdbd7-7aeb-4a13-9e30-797a60057217)


Código para la validacion de datos que son ingresados por el usuario desde la interfaz de angular y se verifica si existe el user y su password en la base de datos, retorna un status OK si el usuario es encontrado y un mensaje de 'Contraseña correcta.
![ValidacionDEUsuarioExistenteEnLaBD](https://github.com/user-attachments/assets/4dfe2242-b331-46a1-9530-04ac7c43605c)


Usando el método get para mostrar los usuarios guardados en la base de datos, ejecutado desde el navegador en la ruta: ```http://localhost:3000/users```.
![MostrandoUsuariosDeLaBD](https://github.com/user-attachments/assets/085a1d18-d15a-4ad3-997d-e95ad220af86)


Código de los métodos Get de 'movies' y 'series' para extraer sus datos desde la base de datos.
![CodigoMetodoGetMoviesSeries](https://github.com/user-attachments/assets/da868ecf-e446-43c3-aae6-eb7f165353eb)


Usando el método get para mostrar el listado de movies en la base de datos, ejecutado desde el navegador en la ruta: ```http://localhost:3000/movies```.
![MostrandoMovies](https://github.com/user-attachments/assets/a743d90e-9304-4168-8258-aabc62e506d9)


Diagrama Entidad-Relación de la base de datos.  
**NOTA: El archivo de la base de datos se encuentra en este mismo repositorio en la carpeta 'DB SQL Server', [DB_SQL_Server](DB_SQL_Server)**
![Diagrama E-R Hub entretenimiento](https://github.com/user-attachments/assets/3c63924d-c57f-4b29-a476-79e87671f9df)


En la siguiente imagen se muestra la página login de la interfaz en Angular conectada a la API en la ruta: ```http://localhost:3000``` y recibiendo respuesta con status: 200, despues de validar que el usuario y la contraseña ingresadas son correctas y existen en la base de datos SQL Server.
![RespuestaDesdeBackend](https://github.com/user-attachments/assets/9f88f3d2-8aa5-4c0a-ab0c-f7f25893db31)


Si el usuario y contraseña son correctas devuelve mensaje 'Sesión iniciada correctamente'
![SesionIniciadaCorrectamente](https://github.com/user-attachments/assets/3b593c7b-9caf-4e8b-ab9c-a820f427e13b)


Si el usuario ingresado no existe en la base de datos se devuelve un error y un mensaje de 'Usuario no registrado'.
![UsuarioNoRegistrado](https://github.com/user-attachments/assets/2f6e306f-8946-49b9-bf0a-d58be51c65bf)


Si el usuario sí existe en la BD pero la contraseña no coincide con la registrada, se devuelve un error y un mensaje de 'Contraseña incorrecta'.
![ContraseñaIngresadaIncorrectamente](https://github.com/user-attachments/assets/abca7b8c-c09c-4cdb-8ecc-5a5dd7a07a5e)


## 5. Proceso de desarrollo

### Detalles
Visual Basic 6.0 fue una tecnología ampliamente utilizada en su tiempo, y aún hoy existen numerosas aplicaciones críticas desarrolladas con este lenguaje. Su relevancia radica en la necesidad de mantener y actualizar estas aplicaciones, que siguen siendo fundamentales para muchas instituciones. Conocer y continuar usando Visual Basic 6.0 es crucial para garantizar la estabilidad y funcionalidad de estos sistemas legacy, lo que subraya la importancia de dominar esta tecnología en el contexto actual.


## 6. Tabla con Sprint Review
**¿Qué salio bien?**  
- Las peticiones en la comunicación con la base de datos para leer, insertar, actualizar y borrar datos funcionó perfectamente.

**¿Qué puedo hacer diferente?**  
- Se pudiera organizar mejor el código, creando funciones, modulos y metodos que realizen ciertas tareas repetitivas para así evitar repeticion del mismo, ademas de hacerlo mas entendible y escalable.


**¿Qué no salio bien?**  
- Me fallo el implementar excepciones y manejo de errores en la aplicacion, para así informar tanto al usuario como al desarrollador de los poblemas que pudieran suceder y como resolverlos.
- Falto implementar mejores alertas para informar al usuario de lo que esta sucediendo en la aplicación y posibles errores.



This project was generated with VisualBasic6.0
