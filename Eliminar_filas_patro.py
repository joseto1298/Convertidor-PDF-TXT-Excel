import pandas as pd
import os

# Obtener la ruta completa del directorio donde se encuentra el archivo de script actual
script_directory = os.path.dirname(os.path.abspath(__file__))

# Obtener una lista de archivos Excel en el mismo directorio
excel_files = [file for file in os.listdir(script_directory) if file.endswith('.xlsx')]

# Lista de patrones que deseas buscar en las filas de cada archivo Excel
patrones = ['- -', 'Ex p']

# Iterar a través de los archivos Excel
for excel_file in excel_files:
    print(f"Procesando archivo: {excel_file}")  # Imprimir el nombre del archivo
    
    # Construir la ruta completa al archivo Excel
    excel_file_path = os.path.join(script_directory, excel_file)
    
    # Leer el archivo Excel
    df = pd.read_excel(excel_file_path)
    
    # Verificar si alguna fila contiene uno de los patrones en las columnas especificadas
    filas_a_eliminar = df.apply(lambda row: any(patron in row.values for patron in patrones), axis=1)
    
    # Eliminar las filas que cumplen con la condición
    df = df[~filas_a_eliminar]
    
    # Actualizar la serie booleana después de eliminar las filas
    filas_a_eliminar = df.apply(lambda row: any(patron in row.values for patron in patrones), axis=1)
    
    # Guardar el DataFrame modificado en el mismo archivo Excel
    df.to_excel(excel_file_path, index=False)
    
    print(f"Filas eliminadas en {excel_file}: {', '.join(map(str, df.index[filas_a_eliminar].tolist()))}")
