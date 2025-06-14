# Exportar archivos en PDFs separados: 

## Nombre del macro: 
 
ImprimirPdfSeparadosv4
 
(No usar V3, es solo un respaldo ya que no se han hecho suficientes pruebas a la v4.
 
La v3 no tiene controles para cancelación que la 4 sí tiene)
 
## Funcionalidad:
 
Permite exportar varios documentos pdf a partir de un solo documento de Word, este funciona bien para documentos que funcionan por el método “Correspondencia” que tiene como DB un documento de Excel. 
 
## Condiciones: 
 
1. Cada archivo debe tener la misma cantidad de páginas
 
2. Ten listo un documento del tipo “txt” que contenga el listado de todos los nombres que quieres que tengan los documentos.
 
    a. Debe ser necesariamente un archivo “.txt”

    b. Debe tener la misma cantidad de nombres que de documentos pdf finales

    c. Deben estar en el mismo orden de nombre que en el Word de página

        i.    Ejemplo:
                - Documento 1 – Nombre 1
                - Documento 2 – Nombre 2
                El macro no tiene forma de identificar automáticamente el contenido de cada página para asignarle el nombre, por lo que no colocar los nombre en el orden correcto terminará en un resultado indeseado.

    d. Los nombres en el documentos deben estar en formato de lista

        i. Ejemplo:
            - Juan Pérez
            - María López
            - Carlos Fernández
            - Ana Torres
            - Luis García
            - Elena Martínez
            - Javier Jiménez
            - Laura Sánchez
            - David Morales
            - Paula Díaz
            - Raúl Castro
            - Teresa Ruiz
            - Sergio Romero
            - Clara Ortega
            - Hugo Navarro
            - Gabriela Herrera
            - Fernando Soto

e. No te preocupes por lo acentos, al macro esta diseñado para poder soportar los acentos correctamente (UTF-8)

3. Sugerentemente ten lista una carpeta en la que deseas tener todos los archivos pdf finales, y en caso de ser muchos, es mejor que los tengas en un lugar.

4. El macro te hará preguntas para los cuales debes tener respuesta:

    a. Cuantas páginas tiene cada uno de los documentos (te recuerdo que deben ser la misma cantidad de páginas para todos los documentos)

    b. Carpeta de exportación (abrirá una venta del explorador de archivos para seleccionarla)

    c. ¿Deseas proporcionar nombre personalizados para los # archivos?

        i. En caso de que sí, ve a “d” de esta lista

        ii. En caso no, te pedirá un nombre base que tendrán todos los archivos, seguido de una numeración automática.

    d. Seleccione el método para proporcionar los nombres:

        iii. Sí, Cargar desde archivo .txt (recomendado)

            Al seleccionar el documento, verificará que contenga la misma cantidad de nombres que de archivos, en caso contrario marcará error y exportará los documentos con el nombre “documento\_(numeración automática)”

        ii. No, Ingresar manualmente

5. Abrirá un cuadro de dialogo que le pedirá ingresar los nombres uno por uno.

6. En caso de que no selecciones nada en caso contrario marcará error y exportará los documentos con el nombre “documento\_(numeración automática)”