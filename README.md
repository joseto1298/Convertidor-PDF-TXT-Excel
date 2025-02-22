# Convertidor de PDF-TXT a Excel

Este proyecto permite extraer tablas de archivos PDF y convertirlas en archivos Excel (.xlsx). Utiliza la librería `tabula` para la extracción de datos y `pandas` para procesarlos.

## Características
- Convierte automáticamente todos los archivos PDF en la carpeta donde se ejecuta el script.
- Soporte para configuración personalizada de columnas según el formato del PDF.
- Filtrado de filas innecesarias y corrección de caracteres especiales.
- Generación de un archivo Excel combinado con los datos de todos los PDFs procesados.

## Requisitos
Asegúrate de tener instaladas las siguientes dependencias antes de ejecutar el script:

```bash
pip install pandas tabula-py openpyxl
```

## Uso
1. Coloca los archivos PDF en la misma carpeta que el script.
2. Ejecuta el script con Python:

```bash
python convert_pdf_to_excel.py
```

3. Se generará un archivo Excel por cada PDF procesado, además de un archivo combinado (`combined.xlsx`) con todos los datos.

## Configuración de Columnas
El script incluye una configuración de columnas personalizada para diferentes tipos de archivos PDF. Puedes modificar la sección correspondiente en el código para ajustarla a tu documento específico.

## Contribuciones
Si deseas mejorar el proyecto, ¡las contribuciones son bienvenidas! Puedes hacer un fork del repositorio y enviar un pull request con tus mejoras.

## Licencia
Este proyecto está bajo la licencia MIT. Puedes usarlo y modificarlo libremente.

