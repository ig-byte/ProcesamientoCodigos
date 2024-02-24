import shutil
import os
import csv
import openpyxl
import re
from tkinter import filedialog

# Definiciones
patronBarcode_regex = re.compile(r'^\d{6}[A-Z]\d{7}$')
indicesReporte = ["Archivo_Origen", "Archivo_Renombrado","Fecha", "Usuario", "Zona", "Linea", "Tanda","Cantidad_Codigos","Inicio/Repaso"]
listaUsuarios = {
    "02": "KCOSMING",
    "03": "MVARAS",
    "04": "ESUAREZ",
    "05": "JROJAS",
    "37": "TEMP1",
    "38": "TEMP2"
}
patronUsuario_regex = re.compile(r'\-(\d{2})\-')
patronLinea_regex = re.compile(r'(LÍNEA|REPASO|LINEA|INICIO|FILA)(\d+)')

# Archivo excel
libro_resultado = openpyxl.Workbook()
hoja_resultado = libro_resultado.active
hoja_resultado.append(indicesReporte)

# Directorios
carpeta_archivos = filedialog.askdirectory(title="Carpeta de destino Fecha")    # Ruta de carpeta con fecha que contiene las listas del día.
archivoReporte = carpeta_archivos.split("/")[-1]                                # Guarda el nombre de fecha de la carpeta de donde se extraen los archivos                            
ruta_original = carpeta_archivos                                                # Directorio de la carpeta con fecha
carpeta_destino = filedialog.askdirectory(title="Carpeta de destino Fecha - R") # Ruta de Carpeta con fecha con sufijo R en donde se guardarán los nuevos excel

# Extrae la fecha del archivo
fecha = carpeta_archivos.split("/")[-1]

# Listado de archivos
archivos_origen = os.listdir(carpeta_archivos) 

# Iterar sobre cada archivo en la carpeta
for archivo_origen in archivos_origen:
    ruta_completa = os.path.join(carpeta_archivos, archivo_origen)
    

    """
    # Verificar si es un archivo CSV
    if archivo_origen.endswith('.csv'):
        # Leer el segundo valor de la segunda fila en el archivo CSV
        with open(ruta_completa, 'r', newline='', encoding='utf-8') as archivo_csv:
            lector_csv = csv.reader(archivo_csv)
            next(lector_csv)  # Ignorar la primera fila
            try:
                valor = next(lector_csv)[1]
                print(f"Archivo CSV: {archivo}, Valor extraído de la segunda fila: {valor}")
            except (StopIteration, IndexError):
                print(f"Archivo CSV: {archivo}, No se encontró información en la segunda fila")
    """
    # Verificar si es un archivo Excel
    if archivo_origen.endswith('.xlsx'):
        # Leer la celda A2 en el archivo Excel
        try:
            libro_excel = openpyxl.load_workbook(ruta_completa)
            hoja_excel = libro_excel.active
            valor = hoja_excel['A2'].value

            # Extrae Usuario
            usuario_temp = re.search(patronUsuario_regex, archivo_origen).group(1)
            if usuario_temp in listaUsuarios:
                usuario = listaUsuarios[usuario_temp]
                idUsuario = archivo_origen.split("-")[1]
               
            else:   #Sin usuario
                print("UXX") 

            # Extracción de string de celda A2
            valor_rev0 = valor.upper()
            valor_rev1 = valor_rev0.replace(" ","")

            # Extraccion de zona
            valor_rev2 = valor_rev1.find("CT")
            zona = valor_rev1[valor_rev2:valor_rev2+4] 

            # Extrae Linea
            coincidenciasLinea = patronLinea_regex.findall(valor_rev1)

            # Iterar sobre las coincidencias y extraer las líneas/repaso
            for coincidencia in coincidenciasLinea:
                palabra_clave, numero_linea = coincidencia
                inicio_indice = valor_rev1.find(coincidencia[0])
                fin_indice = inicio_indice + len(coincidencia[0] + numero_linea)

                valor_rev3_1 = re.findall(r'\d+', valor_rev1[inicio_indice:fin_indice])
                linea = valor_rev3_1[0].zfill(2)


            # Extrae Tanda
            if valor_rev1.rfind("MAÑANA") != -1:
                tanda = "M"
            elif valor_rev1.rfind("TARDE") != -1:
                tanda = "T"
            elif valor_rev1.rfind("M") != -1:
                tanda = "M"
            elif valor_rev1.rfind("T") != -1:
                tanda = "T"
            else:
                tanda = "NoIndicaTanda"
            
            # Cuenta de barcode mediante un patron
            sum = 0
            for celda in hoja_excel['A']:
                if patronBarcode_regex.findall(str(celda.value)):
                        sum+=1
            cantidadCodigos = sum

            #*****************************************************************

            # Nuevo nombre de lista
            nuevo_nombre = f"{fecha}_U{idUsuario}_{zona}_L{linea}_{tanda}_C{cantidadCodigos}"

            print([archivo_origen,nuevo_nombre,fecha, usuario, zona, linea, tanda, cantidadCodigos])
            #shutil.copy(ruta_completa, carpeta_destino)
            #ruta_destino = os.path.join(carpeta_destino, os.path.basename(ruta_completa))
            #os.rename(ruta_destino, os.path.join(carpeta_destino, nuevo_nombre+".xlsx"))



            #hoja_resultado.append([archivo, fecha, usuario, zona, linea, tanda, cantidad de barcode,])

            #print(f"Archivo Excel: {archivo}, - {valor}")
        except Exception as e:
            print(f"Error al procesar el archivo Excel {archivo_origen} - celda {celda.value} - {[archivo_origen,nuevo_nombre,fecha, usuario, zona, linea, tanda, cantidadCodigos]}: {e}")

    # Si es un tipo de archivo desconocido, imprimir un mensaje
    else:
        print(f"Archivo no reconocido: {archivo_origen}")
#libro_resultado.save("Reporte_"+archivoReporte+".xlsx")
#print(f"Archivo Excel creado con la información recolectada: Reporte_{archivoReporte}.xlsx")