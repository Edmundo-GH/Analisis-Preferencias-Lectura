# Análisis de Preferencias de Lectura en Secundaria

Este repositorio contiene el proyecto de inferencia estadística y muestreo realizado a partir de la encuesta de hábitos lectores aplicada a 668 estudiantes de secundaria.

## Contenido del repositorio
* `generar_analisis.py`: Script en Python que procesa los datos, calcula las frecuencias, intervalos de confianza y la prueba Z, y genera un archivo de Excel formateado con gráficas.
* `Base_depurada_normalizada.xlsx`: Base de datos depurada que contiene las respuestas de los alumnos.
* `Analisis_Preferencias_Lectura.xlsx`: (Generado por el script) Libro de Excel con las tablas, cálculos estadísticos y gráficas listas para el informe.

## Requisitos para ejecutar el script
Para correr el código es necesario tener Python instalado y las librerías:
`pip install pandas numpy openpyxl`

## Cómo ejecutar
1. Abrir la terminal en la carpeta del proyecto.
2. Ejecutar el comando: `python generar_analisis.py`
3. Se generará automáticamente el archivo Excel con todos los resultados.