import os
from tkinter import filedialog
import pandas as pd

# Ruta del directorio que contiene los archivos Excel
directorio_archivos = filedialog.askdirectory()
#directorio_archivos = '/ruta/del/directorio'  # Reemplaza con tu ruta

# Lista para almacenar los DataFrames de cada archivo
dataframes = []

# Iterar sobre los archivos en el directorio
for archivo in os.listdir(directorio_archivos):
    if archivo.endswith('.xlsx'):  # Asegúrate de que solo se procesen archivos Excel
        ruta_completa = os.path.join(directorio_archivos, archivo)

        # Leer el archivo Excel y cargarlo como DataFrame
        df = pd.read_excel(ruta_completa)

        # Agregar el nombre del archivo como una nueva columna
        df['Archivo'] = archivo

        # Agregar el DataFrame a la lista
        dataframes.append(df)

# Concatenar todos los DataFrames en uno solo
resultado = pd.concat(dataframes, ignore_index=True)
print(resultado)
resultado.to_excel('resultado1_20230216.xlsx', index=False)

