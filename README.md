# Scraping CONRED

Este proyecto realiza un web scraping básico sobre la página de boletines informativos de CONRED Guatemala. La idea principal es recopilar información pública del sitio, organizarla y generar un archivo de Excel para revisar los datos de una forma más ordenada.

## Objetivo

El objetivo del programa es extraer información de boletines publicados por CONRED, como el título, la clasificación, un resumen y el enlace de la publicación. Esta información puede servir como base para un análisis o para un sistema que necesite consultar comunicados recientes de forma más rápida.

## Herramientas utilizadas

Para realizar el proyecto se utilizó Python junto con algunas librerías sencillas:

- requests: permite conectarse a la página web.
- beautifulsoup4: permite leer y extraer información del HTML.
- openpyxl: permite crear el archivo de Excel con formato.

## Instalación

Antes de ejecutar el programa, se deben instalar las librerías necesarias. Para hacerlo, abra una terminal dentro de la carpeta del proyecto y ejecute:

py -m pip install requests beautifulsoup4 openpyxl

Si el comando anterior no funciona, también puede probar con:

python -m pip install requests beautifulsoup4 openpyxl

## Uso del programa

Para ejecutar el programa, abra la terminal en la carpeta donde se encuentra el archivo scraping_conred.py y escriba:

py scraping_conred.py

Al ejecutarse, el programa revisa varias páginas de boletines informativos de CONRED, extrae la información encontrada y genera automáticamente un archivo llamado:

conred_informacion.xlsx

Este archivo se puede abrir con Microsoft Excel. La información aparece ordenada en columnas, con bordes, texto ajustado y formato para que sea más fácil de leer.

## Datos que se generan

El archivo de Excel contiene las siguientes columnas:

- Título: nombre del boletín o publicación encontrada.
- Clasificación: categoría asignada según el contenido del boletín, por ejemplo actividad volcánica, lluvias, sismos, incendios, prevención o general.
- Resumen: fragmento inicial del contenido encontrado en la publicación.
- Enlace: dirección web del boletín para poder consultarlo directamente.

## Actualización de la información

El archivo de Excel no se actualiza en tiempo real. Para obtener información nueva, se debe cerrar el archivo de Excel y volver a ejecutar el programa.

Cada vez que el programa se ejecuta, consulta nuevamente el sitio web de CONRED y vuelve a generar el archivo conred_informacion.xlsx. Esto permite trabajar con una versión actualizada de los datos.

## Recomendación importante

Si el archivo de Excel está abierto, puede que Python no logre reemplazarlo correctamente. Por eso es recomendable cerrar conred_informacion.xlsx antes de ejecutar nuevamente el programa.

## Nota sobre el uso del scraping

Este proyecto utiliza información pública disponible en el sitio de CONRED. El scraping se realiza con fines académicos y de análisis, evitando el uso de información privada o sensible. También se recomienda no ejecutar el programa de forma excesiva para no realizar demasiadas solicitudes al sitio web.

## Autor

Steve Alvizures.
