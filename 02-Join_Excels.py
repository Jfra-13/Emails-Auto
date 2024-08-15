import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Función para seleccionar la carpeta que contiene los archivos Excel
def seleccionar_carpeta(mensaje):
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de tkinter
    carpeta = filedialog.askdirectory(title=mensaje)
    return carpeta

# Seleccionar la carpeta de origen
ruta_completa = seleccionar_carpeta("Selecciona la carpeta que contiene los archivos Excel")

# Verificar que el directorio existe
if not os.path.exists(ruta_completa):
    print(f"El directorio {ruta_completa} no existe.")
    exit()

# Obtener todos los archivos Excel en el directorio
archivos = [os.path.join(ruta_completa, archivo) for archivo in os.listdir(ruta_completa) if archivo.endswith('.xlsx')]

# Leer cada archivo y guardarlo en un DataFrame
dfs = []
for archivo in archivos:
    try:
        df = pd.read_excel(archivo)
        dfs.append(df)
    except Exception as e:
        print(f"No se pudo leer el archivo {archivo}: {e}")

# Unir todos los DataFrame en uno solo, eliminando filas duplicadas
if dfs:
    resultado = pd.concat(dfs, ignore_index=True)
    
    # Verificar que las columnas importantes existen antes de eliminar duplicados
    columnas_clave = ['DNI', 'Teléfono']
    columnas_existentes = [col for col in columnas_clave if col in resultado.columns]
    
    if len(columnas_existentes) < len(columnas_clave):
        print("Advertencia: No todas las columnas clave están presentes en los archivos. Eliminando duplicados usando las columnas existentes.")
    
    # Eliminar duplicados basados en DNI y Teléfono (si existen)
    resultado = resultado.drop_duplicates(subset=columnas_existentes)
    
    # Eliminar la columna 'N°' si existe y agregar la columna 'ID' con números auto incrementables
    if 'N°' in resultado.columns:
        resultado.drop(columns=['N°'], inplace=True)
    resultado.insert(0, 'ID', range(1, len(resultado) + 1))

    # Seleccionar la carpeta donde se guardará el archivo final
    carpeta_guardado = seleccionar_carpeta("Selecciona la carpeta donde se guardará el archivo de resultado final")

    # Guardar el resultado en un nuevo archivo Excel en la carpeta seleccionada
    ruta_guardado = os.path.join(carpeta_guardado, 'resultadoExcels.xlsx')
    resultado.to_excel(ruta_guardado, index=False)
    print(f"El archivo 'resultadoExcels.xlsx' ha sido creado exitosamente en {carpeta_guardado}.")
else:
    print("No se encontraron archivos para procesar.")

