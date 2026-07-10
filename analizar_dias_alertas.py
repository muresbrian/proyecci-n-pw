import pandas as pd
import os
import glob
from datetime import datetime, timedelta

def main():
    print("Iniciando análisis de días de transacciones...")
    
    # 1. Analizar el histórico de transacciones
    trx_file = 'TRX WU_BP.xlsx'
    print(f"Cargando {trx_file}...")
    try:
        df_trx = pd.read_excel(trx_file, sheet_name='Export')
    except Exception as e:
        print(f"Error cargando TRX: {e}")
        return

    df_trx['Date'] = pd.to_datetime(df_trx['Date'], errors='coerce')
    df_trx = df_trx.dropna(subset=['Date', 'Holder'])

    if 'Abonos' in df_trx.columns:
        df_trx_pos = df_trx[df_trx['Abonos'] > 0].copy()
    else:
        df_trx_pos = df_trx.copy()
    df_trx_pos['DayOfWeek'] = df_trx_pos['Date'].dt.dayofweek
    df_trx_pos['Date_Only'] = df_trx_pos['Date'].dt.normalize()
    
    # Calcular frecuencia relativa de días usando semanas activas
    # Cuantos días únicos con transacciones tiene cada Holder por día de la semana
    trx_por_dia = df_trx_pos.groupby(['Holder', 'DayOfWeek'])['Date_Only'].nunique().unstack(fill_value=0)
    for col in range(7):
        if col not in trx_por_dia.columns:
            trx_por_dia[col] = 0
    # Asegurar orden
    trx_por_dia = trx_por_dia[[0, 1, 2, 3, 4, 5, 6]]
            
    # Calcular semanas activas para cada Holder (fecha_máx - fecha_mín)
    max_date = df_trx_pos['Date_Only'].max()
    first_dates = df_trx_pos.groupby('Holder')['Date_Only'].min()
    total_active_days = df_trx_pos.groupby('Holder')['Date_Only'].nunique()
    weeks_active = ((max_date - first_dates).dt.days / 7.0).apply(lambda x: max(1.0, x))
    
    # Frecuencia en escala de 0 a 100 (% de semanas que registran actividad ese día)
    pct_dias = trx_por_dia.div(weeks_active, axis=0) * 100
    
    # Guardar total de días activos para identificar intermitencia
    pct_dias['Total_Active_Days'] = total_active_days
    
    # Renombrar columnas
    nombres_dias = {0: '%_Lunes', 1: '%_Martes', 2: '%_Miercoles', 3: '%_Jueves', 4: '%_Viernes', 5: '%_Sabado', 6: '%_Domingo'}
    pct_dias = pct_dias.rename(columns=nombres_dias)
    
    # Calcular promedio de Abonos por día
    if 'Abonos' in df_trx_pos.columns:
        promedio_por_dia = df_trx_pos.groupby(['Holder', 'DayOfWeek'])['Abonos'].mean().unstack(fill_value=0)
        # Asegurar que todas las columnas existan
        for col in range(7):
            if col not in promedio_por_dia.columns:
                promedio_por_dia[col] = 0
                
        nombres_promedios = {0: 'Prom_Lunes', 1: 'Prom_Martes', 2: 'Prom_Miercoles', 3: 'Prom_Jueves', 4: 'Prom_Viernes', 5: 'Prom_Sabado', 6: 'Prom_Domingo'}
        promedio_por_dia = promedio_por_dia.rename(columns=nombres_promedios)
        
        # Unir pct_dias con promedio_por_dia
        pct_dias = pct_dias.join(promedio_por_dia)

    
    # 2. Cargar el último reporte de alertas
    alertas_dir = 'REPORTES ALETRAS'
    archivos_alertas = glob.glob(os.path.join(alertas_dir, 'Reporte_Alertas_*.xlsx'))
    # Ignorar los que ya fueron analizados
    archivos_alertas = [f for f in archivos_alertas if not f.endswith('_Analizado.xlsx')]
    
    if not archivos_alertas:
        print("No se encontraron reportes de alertas.")
        return
        
    # Obtener el más reciente
    ultimo_reporte = max(archivos_alertas, key=os.path.getmtime)
    print(f"Procesando reporte de alertas: {ultimo_reporte}")
    
    # Extraer fecha del nombre del archivo (ej. Reporte_Alertas_20260630.xlsx)
    nombre_base = os.path.basename(ultimo_reporte)
    fecha_str = nombre_base.replace('Reporte_Alertas_', '').replace('.xlsx', '').split(' ')[0]
    
    try:
        fecha_reporte = datetime.strptime(fecha_str, '%Y%m%d')
    except:
        fecha_reporte = datetime.now() # Fallback a hoy
        print(f"No se pudo extraer la fecha del archivo {nombre_base}, usando fecha actual.")

    df_alertas = pd.read_excel(ultimo_reporte)
    
    # 3. Cruzar datos
    df_resultado = df_alertas.merge(pct_dias, on='Holder', how='left')
    
    # 4. Determinar si es comportamiento normal
    def es_comportamiento_normal(row):
        dias_sin_trx = row['Dias_sin_transaccionar']
        if pd.isna(dias_sin_trx) or dias_sin_trx <= 0:
            return "No aplica"
            
        # Si no tenemos histórico
        if pd.isna(row['%_Lunes']):
            return "Sin histórico"
            
        dias_sin_trx = int(dias_sin_trx)
        total_trxs = row.get('Total_Active_Days', 0)
        
        # Si el cliente tiene muy pocas transacciones en total en el histórico (< 15 días activos),
        # se clasifica como un cliente intermitente. Es 100% normal que pase días sin transacciones.
        if total_trxs < 15:
            return "Sí, es normal (Cliente intermitente)"
            
        dias_faltantes_nombres = []
        normal = True
        
        # Revisamos los últimos N días
        # Ajustamos el timedelta(days=i-1) para corregir el desfase de 1 día (el día sin transacciones es hoy, no ayer)
        for i in range(1, dias_sin_trx + 1):
            dia_evaluado = fecha_reporte - timedelta(days=i-1)
            dia_semana = dia_evaluado.weekday() # 0 = Lunes, 6 = Domingo
            
            # Nombre de la columna
            col_name = nombres_dias[dia_semana]
            frecuencia = row[col_name]
            
            dias_faltantes_nombres.append(col_name.replace('%_', ''))
            
            # Si transacciona en ese día de la semana en menos del 40% de las semanas activas,
            # consideramos que es normal que falte. Solo si es > 40% indicamos que es inusual.
            if frecuencia > 40:
                normal = False
                
        dias_str = ", ".join(dias_faltantes_nombres)
        if normal:
            return f"Sí, es normal (Faltó {dias_str} donde suele tener baja act.)"
        else:
            return f"No, inusual (Faltó {dias_str} donde suele tener act.)"

    df_resultado['Falta_Normal'] = df_resultado.apply(es_comportamiento_normal, axis=1)
    
    # Reordenar las columnas para mejor legibilidad
    cols = df_resultado.columns.tolist()
    # Poner 'Falta_Normal' justo después de 'Dias_sin_transaccionar' si existe
    if 'Dias_sin_transaccionar' in cols:
        idx = cols.index('Dias_sin_transaccionar')
        cols.insert(idx + 1, cols.pop(cols.index('Falta_Normal')))
        df_resultado = df_resultado[cols]

    # Guardar
    output_path = ultimo_reporte.replace('.xlsx', '_Analizado.xlsx')
    df_resultado.to_excel(output_path, index=False)
    print(f"\n¡Análisis completado! Archivo guardado en: {output_path}")

if __name__ == '__main__':
    main()
