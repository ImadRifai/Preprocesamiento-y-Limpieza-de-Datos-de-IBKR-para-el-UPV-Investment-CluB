# Data Preprocessing - UPV Investment Club 

## Descripción

Este proyecto tiene como objetivo crear un Excel estructurado y fácil de manipular a partir de los datos de nuestra Cartera descargados de IBKR, útil para diversos análisis (por ej. el dashboard en Power BI para el Club). El código está diseñado para ser sencillo, comentado y modular, permitiendo así modificaciones de otros miembros en un futuro.

## Objetivo

- Crear un archivo Excel estructurado a partir del informe desordenado y sin formato de IBKR.
- Obtener tablas que tengan la mayor cantidad de información relevante del informe original.
- Escribir un código claro y bien comentado, para que cualquier miembro del club pueda hacer modificaciones en el futuro.

## Instrucciones de Uso

1. Coloca el archivo de IBKR en formato Excel en la misma carpeta que este script. Y cambia el nombre_report, al nombre del archivo de IBKR.
2. Ejecuta el script en Python.
3. Obtendrás el archivo de salida en la misma carpeta**

## Dependencias

Este script requiere la librería Pandas de Python:

Instala las dependencias ejecutando:
```bash
pip3 install pandas
```

## Archivos de Ejemplo
En el repositorio, se incluyen:

*  Excel de Entrada: es el archivo descargado de IBKR en su estado original (desordenado y sin formato), con los datos anonimizados.
*  Excel de Salida: es el resultado final del script, en este ejemplo los datos han sido anonimizados, pero se puede observar la estructura final y se refleja la transformación de un formato crudo a uno organizado.

## Dashboard con PowerBI

Tras la obtención, limpieza y procesamiento de los datos, como parte posterior del proyecto he creado el dashboard del club. 
Se puede observar el dashboard aquí:

https://app.powerbi.com/view?r=eyJrIjoiYzQ0NGEzYmUtNjM4Zi00ODk0LWE5NjYtNTRkMDAwMGEzYWE2IiwidCI6ImJlNDY1NWRmLWFjNzMtNDAxZi1hN2FlLTE5OGMzYjcyZDBjNiIsImMiOjh9
