import pandas as pd
from tkinter import filedialog

# Seleccionar el archivo principal que contiene los datos existentes
archivo_principal = filedialog.askopenfilename(title="Seleccionar archivo principal", filetypes=[("Archivos Excel", "*.xlsx")])
df_principal = pd.read_excel(archivo_principal)

# Seleccionar el archivo que contiene los campos adicionales
archivo_campos_adicionales = filedialog.askopenfilename(title="Seleccionar archivo con campos adicionales", filetypes=[("Archivos Excel", "*.xlsx")])
df_campos_adicionales = pd.read_excel(archivo_campos_adicionales)

# Fusionar los DataFrames en base al campo común 'Archivo'
resultado_final = pd.merge(df_principal, df_campos_adicionales, on='Archivo', how='left')

# Guardar el resultado en un nuevo archivo Excel
test = filedialog.askdirectory()
resultado_final.to_excel(test+'/'+'resultado_final_20240216.xlsx', index=False)

print("Campos agregados correctamente. Resultado guardado en 'resultado_final_rev2.xlsx'.")
