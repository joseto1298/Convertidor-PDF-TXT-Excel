import os
import pandas as pd

def skip_rows_with_keywords(line, keywords):
    # Verificar si la línea contiene las palabras clave de manera consecutiva
    for keyword in keywords:
        if keyword in line:
            return True
    return False

def convert_txt_to_excel(txt_file, output_file, column_positions, skip_keywords):
    try:
        # Leer el archivo de texto y dividirlo en líneas
        with open(txt_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        # Crear una lista para almacenar las filas de datos
        data_rows = []

        # Iterar a través de las líneas del archivo de texto
        skip_line = False  # Bandera para omitir filas
        for line_index, line in enumerate(lines):
            # Verificar si la fila contiene palabras clave de manera consecutiva
            if skip_line or skip_rows_with_keywords(line, skip_keywords):
                skip_line = True
                continue
            
            # Si no es la primera línea (encabezado), verificar las palabras clave
            if line_index > 0:
                if "CMAÑOEXP " in line or "Tarifa    " in line:
                    continue
                        
            # Extraer datos de cada columna según las posiciones especificadas
            columns = [line[pos[0]:pos[1]].strip() for pos in column_positions]

            # Realizar las sustituciones necesarias en las columnas
            columns = [col.replace('Ã±', 'ñ').replace('Ã‘', 'Ñ').replace('l ', 'l').replace('Nº ', 'Nº') for col in columns]

            #columns[4] = columns[4].replace(',', '').replace('.', '')    
            #columns[12] = columns[20].replace(',', '').replace('.', '')
            columns[10] = columns[10].replace(',', '').replace('.', '')
            columns[11] = columns[11].replace(',', '').replace('.', '')
            columns[13] = columns[13].replace(',', '').replace('.', '')
            columns[14] = columns[14].replace(',', '').replace('.', '')

            #columns[4] = columns[4].zfill(5)
            #columns[20] = columns[20].zfill(5)
            columns[10] = columns[10].zfill(5)
            columns[11] = columns[11].zfill(5)
            columns[13] = columns[13].zfill(5)
            columns[14] = columns[14].zfill(5)

            #columns[11] = columns[11].replace(',', '').replace('.', '')  # Columna 10
            #columns[12] = columns[12].replace(',', '').replace('.', '')  # Columna 11

            # Insertar guiones en la columna 5 para obtener el formato "00-00-00"
            #columns[5] = '-'.join([columns[5][:2], columns[5][2:4], columns[5][4:]])


            # Agregar la fila de datos a la lista
            data_rows.append(columns)

        # Crear un DataFrame a partir de las filas de datos
        df = pd.DataFrame(data_rows)

        # Escribir el DataFrame en un archivo de Excel
        df.to_excel(output_file, index=False, header=False)

        print(f"Se ha convertido exitosamente el archivo TXT a Excel. Puedes encontrar el archivo en: {output_file}")
    except Exception as e:
        print(f"Ha ocurrido un error al convertir el archivo TXT a Excel: {str(e)}")

# Obtener la ruta de la carpeta actual donde se encuentra el script
script_folder = os.path.dirname(os.path.abspath(__file__))

# Obtener la lista de archivos TXT en la carpeta
txt_files = [f for f in os.listdir(script_folder) if f.endswith('.txt')]

# Lista para almacenar los DataFrames de los archivos Excel
excel_dfs = []

# Definir las posiciones de caracteres para cada columna
column_positions =  [(0, 9),(9,18),(18,28),(28,65),(65,72),(72,79),(79,87),(87,94),(94,102),(102,135),(135,144),(144,153),(153,186),(186,194),(194,204),(204,211),(211,217),(217,226),(226,234),(234,242),(242,251),(251,267),(267,304),(304,314),(314,600)]  # RSQCEM1

#column_positions = [(0, 9),(9,18),(18,51),(51,60),(60,69),(69,79),(79,116),(116,124),(124,132),(132,139),(139,147),(147,155),(155,162),(162,170),(170,178),(178,186),(186,194),(194,201),(201,208),(208,216),(216,224),(224,232),(232,240),(240,248),(248,256),(256,265),(265,283),(283,293),(293,300),(300,320)] # RSQCEM2

#column_positions = [(0, 9),(9,18),(18,25),(25,101),(101,150)] # RSQCEM3

# Definir palabras clave para omitir filas
skip_keywords = []

# Recorrer cada archivo TXT y convertirlo a Excel
for txt_file in txt_files:
    # Ruta completa del archivo TXT
    txt_path = os.path.join(script_folder, txt_file)

    # Nombre de archivo Excel de salida
    excel_file = os.path.splitext(txt_file)[0] + '.xlsx'

    # Ruta completa del archivo Excel de salida
    output_file = os.path.join(script_folder, excel_file)

    # Llamar a la función para convertir el TXT a Excela
    convert_txt_to_excel(txt_path, output_file, column_positions, skip_keywords)

    # Leer el archivo de Excel recién creado y agregar el DataFrame a la lista
    df_excel = pd.read_excel(output_file, header=None)
    excel_dfs.append(df_excel)

# Concatenar los DataFrames en uno solo
#df_concatenated = pd.concat(excel_dfs)

# Escribir el DataFrame concatenado en un archivo de Excel
#output_combined_file = os.path.join(script_folder, "combined.xlsx")
#df_concatenated.to_excel(output_combined_file, index=False, header=False)

#print(f"Se han combinado exitosamente los archivos Excel. Puedes encontrar el archivo combinado en: {output_combined_file}")
