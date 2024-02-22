import pandas as pd
import os
from tkinter import filedialog
import shutil

# Ruta de la carpeta que contiene los archivos Excel
carpeta_excel = filedialog.askdirectory(title="Carpeta excel")
# Ruta de la carpeta donde se guardarán los archivos CSV
carpeta_csv  = filedialog.askdirectory(title="Carpeta CSV")

# Obtener la lista de archivos Excel en la carpeta
archivos_excel = [archivo for archivo in os.listdir(carpeta_excel) if archivo.endswith('.xlsx') or archivo.endswith('.xls')]

# Iterar sobre los archivos Excel y convertirlos a CSV
for archivo_excel in archivos_excel:
    # Leer el archivo Excel
    df = pd.read_excel(os.path.join(carpeta_excel, archivo_excel))

    # Crear la ruta para el archivo CSV de salida
    nombre_csv = os.path.splitext(archivo_excel)[0] + '.csv'
    ruta_csv = os.path.join(carpeta_csv, nombre_csv)

    # Guardar el DataFrame como un archivo CSV
    df.to_csv(ruta_csv, index=False)

print("Proceso completado. Archivos CSV generados en:", carpeta_csv)
