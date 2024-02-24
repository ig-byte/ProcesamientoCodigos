import pandas as pd
from tkinter import filedialog
import shutil
import os

# Lee el archivo CSV

carpeta_archivos = filedialog.askdirectory(title="Carpeta con CSV") 
archivos_csv = os.listdir(carpeta_archivos)


for archivo in archivos_csv:
    datos = pd.read_csv(os.path.join(carpeta_archivos,archivo), encoding='latin-1')
    # Muestra los primeros registros para verificar la estructura del archivo
    print("Datos originales:")
    print(datos.head())

    # Guarda los datos en un nuevo archivo Excel con columnas separadas
    archivo_excel = archivo.split(".")[0]+'.xlsx'
    datos.to_excel(archivo_excel, index=False)

    # Muestra los primeros registros del nuevo archivo para verificar la transformación
    datos_transformados = pd.read_excel(archivo_excel)
    print("\nDatos transformados:")
    print(datos_transformados.head())