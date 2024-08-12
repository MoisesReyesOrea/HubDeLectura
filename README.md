# Hub de lectura - Visual Basic 6.0, WinForm proyecto para Mega - Liderly

## 1. Descripción
Este repositorio contiene una aplicación de escritorio para windows hecha con el lenguaje Visual Basic 6.0 y SQL Server (Transact SQL). Esta aplicación windows es acerca de un Hub de libros en la cual se pueden agregar usuarios, cada usuario puede seleccionar su contenido como Favoritos, libros ya leidos, libros que está leyendo, libros no deseados, ademas de poder visualizar, es posible añadir, eliminar y modificar tanto usuarios como libros. 

### Objetivo
El objetivo es crear una aplicación windows con visual basic 6.0 en que el usuario pueda organizar y controlar los libros que desea leer.

## 2. Requerimientos técnicos:
### Para visualizar el contenido del codigo es necesario tener instalado:  
Microsoft Visual Basic 6.0.  
GIT: Debe tener Instalado GIT.  
SQL Server: En este caso la aplicación se conecta a una base de datos local realizada en SQL Server.  

## 3. ¿Cómo ejecutar la aplicación?

### Clonando el proyecto y ejecutandolo desde Microsoft Visual Basic 6.0 (Desarrollador)
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

### Ejecuta el archivo HubDeLecturaMega.exe (Usuario)  
--   descarga y ejecuta el archivo HubDeLectura.exe, te instalará la aplicación.  

## 4. Explicación
Ventana de registro para un nuevo usuario, el registro se lleva a cabo en la base de datos para despues ser validado en el inicio de sesión.  
![Signup Creacion de usuario](https://github.com/user-attachments/assets/81b06a35-aabd-4316-aa8c-def18da732df)  

Ventana de login, los datos ingresados se buscan en la base de datos si se encuentran se inicia sesión.
![Inicio de sesion](https://github.com/user-attachments/assets/24fcbcfb-7d64-4b61-8946-8ef8e53ed531)  

Ventana de inicio de la aplicación despues de haber iniciado correctamente sesión, en esta sección se muestra la lista completa de los libros disponibles y tambien se puede añadir un nuevo libro al registro.
![InicioRegistrarNuevoLibro](https://github.com/user-attachments/assets/60e115df-bd25-4af3-9c7e-0054f14f3444)  

Notificación despues de agregar un nuevo libro, se añade a la base de datos como a la lista que muestra todos los libros disponibles.  
![LibroAñadido](https://github.com/user-attachments/assets/60b35cd4-de34-4e4a-ba59-2d52e58152d6)  

Añadiendo un libro a la sección de favoritos, esto se relaciona al id de libro y al id del usuario, se almacena el registro en la BD.
![AñadirLibroAFavoritos](https://github.com/user-attachments/assets/74b5530e-a986-4720-8678-bf453ec48bb9)

Ventana de lista de libros favoritos, aquí tambien se puede eliminar los libros de la lista.  
![VentanaFavoritos](https://github.com/user-attachments/assets/31649901-2412-4c09-ba85-665dfb73acaa)  

Notificación de libro eliminado
![LibroEliminadoDeFavoritos](https://github.com/user-attachments/assets/51323c9f-c109-4227-a13d-5047448c96aa)  

Ventana de libros leidos o completados, tambien se cuenta con las ventanas de libros leyendo y libros no deseados.
![VentanaLibrosCompletados](https://github.com/user-attachments/assets/2dc8c27a-4118-4ea9-90d5-c5c92a4b07a2)  

Sección de perfil de usuario, aquí se muestra toda la información perteneciente al usuario, se puede eliminar o modificar su información.  
![VentanaPerfilUsuario](https://github.com/user-attachments/assets/fdb9c9cf-5c47-47bb-ad1a-048cc140bde8)  

Para la modificación de datos el botón 'Modificar' se habilita despues de alterar cualquiera de los datos en los campos.  
![ModificarInformacionUsuario](https://github.com/user-attachments/assets/39a49115-4031-4e5e-94ed-350d815b2040)  

Al presionar el boton de Eliminar usuario se envia una alerta de confirmación antes de realizar la eliminación desde la BD.
![EliminacionDeUsuario](https://github.com/user-attachments/assets/0a31dceb-ac90-4cde-86d1-76f3e947ae89)  

En la ventana de inicio se puede elegir un libro y selecionar el boton 'Ver libro' para visualizar la información completa del mismo, tambien en esta sección se muestran los botones 'Marcar como leido', 'Agregar a leyendo' o 'Marcar como no deseado' para añadirlo a cualquiera de las listas correspondientes del usuario.  
![VentanaVisualizacionDeLibro](https://github.com/user-attachments/assets/20b61a35-7f0c-4bc5-87b3-6e70ca5df306)  

Ventana de lista de libros leyendo, aquí el usuario lleva el registro de los libros que esta leyendo en ese momento antes de completarlos.  
![VentanaDeListaLibrosLeyendo](https://github.com/user-attachments/assets/c91ed7e3-1dad-4ad1-b5df-c4c91601026a)

Tambien se puede eliminar un libro de la lista general.  
![LibroEliminado](https://github.com/user-attachments/assets/bc66e363-449d-4e5c-89d2-1d5d194e0171)  

Diagrama Entidad Relación de la Base de Datos hecha en SQL Server.  
![Diagrama ER base de datos](https://github.com/user-attachments/assets/0429a004-5f2d-465f-b3cd-ea7f08e41d2b)  

Fragmento de los comandos para la creación de las tablas de la aplicación, **NOTA: Los archivos de la BD estan en este mismo repositorio en la carpeta 'DB_SQL_Server', [DB_SQL_Server](DB_SQL_Server)**  
![Codigo T SQL base de datos](https://github.com/user-attachments/assets/ecef1d89-f651-4932-8ada-fe688c743818)  

## 5. Proceso de desarrollo

### Detalles
Visual Basic 6.0 fue una tecnología ampliamente utilizada en su tiempo, y aún hoy existen numerosas aplicaciones críticas desarrolladas con este lenguaje. Su relevancia radica en la necesidad de mantener y actualizar estas aplicaciones, que siguen siendo fundamentales para muchas instituciones. Conocer y continuar usando Visual Basic 6.0 es crucial para garantizar la estabilidad y funcionalidad de estos sistemas legacy, lo que subraya la importancia de dominar esta tecnología en el contexto actual.


## 6. Tabla con Sprint Review
**¿Qué salio bien?**  
- Las peticiones en la comunicación con la base de datos para leer, insertar, actualizar y borrar datos funcionó perfectamente.

**¿Qué puedo hacer diferente?**  
- Se pudiera organizar mejor el código, creando funciones, modulos y metodos que realizen ciertas tareas repetitivas para así evitar repeticion del mismo, ademas de hacerlo mas entendible y escalable.
- Mejorar la interfaz de usuario, agregar un poco más de estilo a las ventanas de visualización.

**¿Qué no salio bien?**  
- Me fallo el implementar excepciones y manejo de errores en la aplicacion, para así informar tanto al usuario como al desarrollador de los poblemas que pudieran suceder y como resolverlos.
- Falto implementar mejores alertas para informar al usuario de lo que esta sucediendo en la aplicación y posibles errores.



This project was generated with VisualBasic6.0
