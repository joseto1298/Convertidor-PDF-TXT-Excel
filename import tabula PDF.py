import os
import tabula
import pandas as pd

def convert_pdf_to_excel(pdf_file, output_file):
    try:
        # Configuración adicional para mejorar la separación de columnas
        tabula_options = {
            'pages': 'all',
            'guess': False,
            'area': (0, 0, 1000, 1000),  # Ajustar las coordenadas del área según tu PDF
            #/////////////////////////////////////////////////////
            #'columns': [70.1, 110.2,170.3, 320.4, 400.5]  # RSQOBR01
            #'columns': [70.1, 110.2, 1000.3]  # RSQOBR02 RSQOBR03
            #'columns': [70.1, 110.2, 365.3, 420.4, 800.5]  # RSQOBR04
            #'columns': [70.1, 110.2, 395.3, 420.4, 470.5, 540.6, 900.7]  # RSQOBR05
            #'columns': [70.1, 110.2, 175.3, 230.4, 290.5, 380.6, 450.7, 510.8, 900.9]  # RSQOBR06
            #'columns': [70.1, 110.2, 240.3, 330.4, 360.5, 430.6, 900.7]  # RSQOBR07
            #'columns': [70.1, 110.2, 190.3, 900.4]  # RSQOBR08
            #/////////////////////////////////////////////////////
            #'columns': [55.1, 70.2, 100.3, 130.4, 390.5, 470.6, 1000.7]  # RSQRVO01
            #'columns': [55.1, 70.2, 100.3, 130.4, 160.5, 220.6, 1000.7]  # RSQRVO02
            #'columns': [55.1, 70.2, 100.3, 130.4, 470.5, 535.6, 1000.7]  # RSQRVO03
            #'columns': [55.1, 70.2, 100.3, 130.4, 205.5, 280.6, 515.7, 550.8, 1000.9]  # RSQRVO04
            #'columns': [55.1, 70.2, 100.3, 130.4, 195.5, 225.6, 300.7, 350.8, 1000.9]  # RSQRVO05
            #'columns': [55.1, 70.2, 100.3, 130.4, 9000.5,]  # RSQRVO06
            #/////////////////////////////////////////////////////
            #'columns': [55.1, 100.2, 135.3, 175.4, 485.5, 1000.6]  # RSQFVO01
            #'columns': [55.1, 100.2, 140.3, 175.4, 430.5, 485.6, 1000.7]  # RSQFVO02
            #'columns': [55.1, 100.2, 140.3, 175.4, 265.5, 305.6, 1000.7]  # RSQFVO03
            #'columns': [55.1, 100.2, 140.3, 175.4, 1000.7]  # RSQFVO04
            #'columns': [55.1, 100.2, 140.3, 175.4, 1000.7]  # RSQFVO05
            #'columns': [55.1, 100.2, 140.3, 175.4, 250.5, 320.5, 380.6, 430.7, 470.8, 530.9, 10000.10]  # RSQFVO06
            #'columns': [55.1, 100.2, 140.3, 175.4, 430.5 , 1000.6]  # RSQFVO07
            #/////////////////////////////////////////////////////
            #'columns': [55.1, 100.2, 125.3, 215.3, 230.4, 290.5, 1000.6]  # RSQAPR01
            #'columns': [55.1, 100.2, 1000.3]  # RSQAPR02
            #'columns': [55.1, 100.2, 445.3, 1000.3]  # RSQAPR03
            'columns': [55.1, 100.2, 350.3, 390.4, 460.5, 500.6, 1000.7]  # RSQAPR04
            #'columns': [55.1, 100.2, 175.3, 240.4, 305.5, 365.6, 400.7, 465.8, 490.9, 510.10 ,1000.11]  # RSQAPR05
            #'columns': [55.1, 100.2, 1000.3]  # RSQAPR06

        }

        # Extraer tablas del PDF y guardarlas en una lista de DataFrames
        dfs = tabula.read_pdf(pdf_file, **tabula_options)

        # Concatenar los DataFrames en uno solo
        df_concat = pd.concat(dfs)

        # Filtrar filas que contengan guiones consecutivos en cualquier columna del DataFrame
        df_concat = df_concat[~df_concat.astype(str).apply(lambda x: x.str.contains('- - -')).any(axis=1)]
        df_concat = df_concat[~df_concat.astype(str).apply(lambda x: x.str.contains('- - - - - -')).any(axis=1)]
        
        # Realizar las sustituciones en el DataFrame
        df_concat.replace({'Ã±': 'ñ', 'I ': 'I', 'J ': 'J', 'Ã‘': 'Ñ', 'l ': 'l','Nº ': 'Nº' }, inplace=True, regex=True)
        
        #Eliminar comas y espacios en la cuarta columna
        df_concat.iloc[:, 3] = df_concat.iloc[:, 3].str.replace(',', '').str.strip()
        df_concat.iloc[:, 3] = df_concat.iloc[:, 3].str.replace(' ', '').str.strip()

        df_concat.iloc[:, 4] = df_concat.iloc[:, 4].str.replace('.', '').str.strip()
        df_concat.iloc[:, 4] = df_concat.iloc[:, 4].str.replace(' ', '').str.strip()      

        # Agregar ceros a la izquierda si la columna tiene menos dígitos
        df_concat.iloc[:, 4] = df_concat.iloc[:, 4].str.zfill(5)
        df_concat.iloc[:, 3] = df_concat.iloc[:, 3].str.zfill(5)

    
        # Insertar guiones "-" en la cuarta columna para obtener el formato 00-00-00
        #df_concat.iloc[:, 7] = df_concat.iloc[:, 7].str[:2] + '-' + df_concat.iloc[:, 7].str[2:4] + '-' + df_concat.iloc[:, 7].str[4:]
        #df_concat.iloc[:, 10] = df_concat.iloc[:, 10].str[:2] + '-' + df_concat.iloc[:, 10].str[2:4] + '-' + df_concat.iloc[:, 10].str[4:]

        # Escribir el DataFrame en un archivo de Excel
        df_concat.to_excel(output_file, index=False)

        print(f"Se ha convertido exitosamente el archivo PDF a Excel. Puedes encontrar el archivo en: {output_file}")
    except Exception as e:
        print(f"Ha ocurrido un error al convertir el archivo PDF a Excel: {str(e)}")


# Obtener la ruta de la carpeta actual donde se encuentra el script
script_folder = os.path.dirname(os.path.abspath(__file__))

# Obtener la lista de archivos PDF en la carpeta
pdf_files = [f for f in os.listdir(script_folder) if f.endswith('.pdf')]

# Lista para almacenar los DataFrames de los archivos Excel
excel_dfs = []

# Recorrer cada archivo PDF y convertirlo a Excel
for pdf_file in pdf_files:
    # Ruta completa del archivo PDF
    pdf_path = os.path.join(script_folder, pdf_file)

    # Nombre de archivo Excel de salida
    excel_file = os.path.splitext(pdf_file)[0] + '.xlsx'

    # Ruta completa del archivo Excel de salida
    output_file = os.path.join(script_folder, excel_file)

    # Llamar a la función para convertir el PDF a Excel
    convert_pdf_to_excel(pdf_path, output_file)

    # Leer el archivo de Excel recién creado y agregar el DataFrame a la lista
    df_excel = pd.read_excel(output_file)
    excel_dfs.append(df_excel)

# Concatenar los DataFrames en uno solo
df_concatenated = pd.concat(excel_dfs)

# Escribir el DataFrame concatenado en un archivo de Excel
output_combined_file = os.path.join(script_folder, "combined.xlsx")
df_concatenated.to_excel(output_combined_file, index=False)

print(f"Se han combinado exitosamente los archivos Excel. Puedes encontrar el archivo combinado en: {output_combined_file}")
