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
