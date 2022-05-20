# Sistema de Información para la Gestión de Almacén
Este proyecto consta de un sistema de información realizado bajo Visual Basic 6 y con base de datos MS Access para la gestión de los productos dentro de un almacén de una empresa Venezolana.

Todos los archivos necesarios para su funcionamiento se encuentran en la carpeta *Project*. Las demás carpetas que se encuentran en el proyecto son los que contienen los archivos necesarios para manipular el código.

## Requisitos

- Sistema operativo Windows 7 o posterior. (Testeado en Windows 10 y 7).
- Microsoft Office 2007 o posterior con Access.


## Instalación
Antes de poder probar el proyecto se deben de instalar/registrar los siguientes archivos DLL.

- Ir a la carpeta Project/DLL y copiar todos los archivos que se encuentran ahí.
- Pegarlos en la siguiente ruta:
    - Para windows de 32 bits

        `C:\WINDOWS\System32\`

    - Para windows de 64 bits

        `C:\WINDOWS\SysWOW64\`
- Una vez pegados se procede a registrar **cada archivo** con la consola/cmd/símbolo de sistema de windows **(se recomienda hacerlo en modo administrador)** de la siguiente manera:
    - Para windows de 32 bits:

        `regsvr32.exe C:\WINDOWS\System32\NOMBRE_ARCHIVO.OCX`

    - Para windows de 64 bits:

        `regsvr32.exe C:\WINDOWS\SysWOW64\NOMBRE_ARCHIVO.OCX`

## Usar el proyecto
Una vez instalado el proyecto se puede probar el proyecto. El mismo se encuentra en su versión ya empaquetada (ejecutable) dentro de la carpeta Project con el nombre de *GestionAlmacen.exe*.

El usuario por defecto es: `administrador`

La contraseña por defecto es `Administrador`

**IMPORTANTE:** Si se quiere trasladar el proyecto a otra ruta, se debe llevar las carpetas:

- Backup
- ImagenesProductos
- Interfaz
- Manual
- Y la base de datos llamada BDProyecto.mdb