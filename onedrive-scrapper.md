# OneDrive Scrapper

Quiero descargar ficheros de forma automatizada de unas carpetas de OneDrive.

Como no tengo acceso a la API de OneDrive, quiero implementar un web scrapper con las siguientes características:

* Se implementará como una librería Node, de forma que pueda usarse en varios proyectos Node que tenga la necesidad de interactuar con OneDrive.
* Como la autenticación a la carpeta requiere multi-factor, al usuario se le mostrará un navegador en el que se tenga que autenticar.
* El navegador usado se persistirá para que el usuario no tenga que autenticarse cada vez que se ejecute el web scrapper.
* De momento el scrapper sólo permitirá la lectura de carpetas y descarga de ficheros. Aunque debe estar preparado para subida de ficheros y compartición en el futuro.
* Para que el coding agent aprenda la forma de interactuar con la página, se le proporciona una carpeta con ficheros bien conocidos:
  * https://urjc-my.sharepoint.com/my?id=%2Fpersonal%2Fmicael%5Fgallego%5Furjc%5Fes%2FDocuments%2FPruebaOneDriveScrapper
  * Contenido:
    * Doc1.docx
    * Doc2.xlsx
    * Folder/Doc3.pptx
    * Folder/Doc4.md
* La API de la librería permitirá explorar las carpetas y descargar el contenido de los ficheros de forma similar a como operan las librerías de acceso al sistema de ficheros.
* La API de la librería permitirá las siguientes acciones de alto nivel:
  * Descargar la estructura de subcarpetas y nombres de ficheros de una carpeta determinada en txt y json.
  * En la descarga se podrá configurar que reconstruya la estructura de carpetas y ficheros en el disco local.
  * Poner filtros en el nombre de los ficheros que se descargan. El filtro puede ser del nombre del fichero o de la ruta completa desde la carpeta "raíz".
  * Poner filtros en el nombre de las carpetas que se exploran. El filtro puede ser del nombre del fichero o de la ruta completa desde la carpeta "raíz".
* Las acciones de alto nivel también se proporcionarán en un CLI que pueda ser usado desde la línea de comandos.
* Se deberá crear un documento describiendo tanto la API de la librería (con ejemplos de código de uso) como los comandos del CLI.