# J3Cnn-Suite
Este es un conjunto de 3 dll’s Activex que permiten manipular base de datos de SQLite3, bases de datos cifradas de SQLCipher y bases de datos cifradas de SQLite3MultipleCiphers, no requiere de compilaciones personalizadas, se incluye el proyecto de ejemplo sobre su uso, básico y avanzado, tanto para base de datos cifrada y no cifrada, se incluye el módulo J3cnnLoader.bas que permite usar cualquiera de los tres componentes sin la necesidad de registrar la DLL ActiveX.

Para más información vea el archivo [Readme-es](readme-es.pdf)

## USO
Para usar cualquiera de estos componentes sin registrarlos en el sistema, añada la referencia a la dll e incluya las dlls requeridas en la carpeta o sub carpeta de su proyecto y añada a su proyecto el módulo J3cnnLoader.bas y listo.

También puede cargar cualquiera de los tres componentes desde una ruta personalizada con la API LoadLibrary() y el módulo se encargará de crear los objetos necesarios. En el módulo J3cnnLoader.bas existe la función LoadLib() que es un envoltorio que permite cargar librerías desde rutas personalizadas.


