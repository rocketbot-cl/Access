



# Microsoft Access
  
Automate Microsoft Access  
  
![banner](imgs/Banner_Access.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  




## Como usar este módulo

Para usar este módulo, tienes que seleccionar una base de datos (.mdb o .accdb) y 
connectarte a ella; luego puedes ejecutar query para obtener los datos de la misma.
Para leer base de datos .accdb, 
deberá tener instalado los respectivos drivers (se deben instalar los drivers de 32 bits).
Se pueden bajar desde el 
siguiente link:
https://www.microsoft.com/es-es/download/details.aspx?id=13255




## Descripción de los comandos

### Conectarse a Base de Datos
  
Comando para conectarse a una base de datos desde un archivo access
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Seleccione base de datos|Conecta a la base de datos creada anteriormente. Formatos .mdb y .accdb son aceptados.|C:\Ruta\a\basededatos.accdb|

### Ejecutar query
  
Ejecuta query en una base de datos
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Query|Query que se desea ejecutar en la base de datos.|select * from tabla|
|Asignar resultado a variable|Variable donde guardar el resultado.|Variable|

### Cerrar conexion
  
Cierra la conexión establecido a la base de datos de Access.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
