import os
import pandas as pd

# Rutas de tus archivos (relativas a la carpeta contenedora superior)
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
archivo_txt = os.path.join(parent_dir, "PRUEBa hob.txt")
archivo_excel = os.path.join(parent_dir, "PRUEBa hob.xlsx")

try:
    # Leemos el archivo de texto. 
    # El parámetro 'sep' indica cómo están separados los datos. 
    # '\t' significa que están separados por una tabulación (muy común al copiar/pegar).
    # Si tus datos están separados por comas, cámbialo a sep=',' o por punto y coma sep=';'
    df = pd.read_csv(archivo_txt, sep='\t')
    
    # Exportamos los datos a un archivo de Excel
    df.to_excel(archivo_excel, index=False, engine='openpyxl')
    
    print(f"¡Éxito! El archivo se guardó como Excel en:\n{archivo_excel}")

except Exception as e:
    print(f"Ocurrió un error al intentar convertir el archivo: {e}")