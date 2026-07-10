import pandas as pd
import os

def separar_hojas_csv():
    archivo_origen = 'Analisis_Historico_Variaciones.xlsx'
    
    # Validar que el archivo maestro exista
    if not os.path.exists(archivo_origen):
        print(f"[ERROR] No se encontro el archivo '{archivo_origen}'. Asegurate de haber corrido el analisis principal primero.")
        return

    print(f"Leyendo el archivo: {archivo_origen}...")
    
    # Leer todas las hojas
    hojas = pd.read_excel(archivo_origen, sheet_name=None)
    
    # Crear una carpeta para los CSV
    carpeta_salida = os.path.join('webapp', 'Reportes_Individuales_CSV')
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
        
    print(f"Separando {len(hojas)} hojas y guardando como CSV...\n")
    
    # Guardar cada hoja como un CSV independiente
    for nombre_hoja, df_hoja in hojas.items():
        archivo_destino = os.path.join(carpeta_salida, f"{nombre_hoja}.csv")
        
        # Guardamos en CSV con utf-8-sig para que lea bien los emojis y acentos
        df_hoja.to_csv(archivo_destino, index=False, encoding='utf-8-sig')
        print(f"[OK] Archivo generado: {nombre_hoja}.csv")
        
    print(f"\n[OK] ¡Proceso completado! Tus archivos estan dentro de la carpeta '{carpeta_salida}'.")

if __name__ == '__main__':
    separar_hojas_csv()