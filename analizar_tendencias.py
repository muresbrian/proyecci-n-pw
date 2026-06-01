import pandas as pd
import numpy as np

def procesar_datos():
    print("Cargando y limpiando datos operativos (2026)...")
    # 1. Cargar hoja principal (Diaria 2026)
    df = pd.read_excel('TRX WU_BP.xlsx', sheet_name=0)
    
    df = df[df['Date'] != 'Total']
    df = df[df['Holder'] != 'Total']
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])

    # Periodos Temporales
    df['Mes'] = df['Date'].dt.to_period('M').astype(str)
    df['Semana'] = df['Date'].dt.to_period('W').astype(str)
    df['Dia'] = df['Date'].dt.date.astype(str)
    df['Quincena'] = df['Date'].dt.day.apply(lambda x: '1ra Quincena (Días 1-15)' if x <= 15 else '2da Quincena (Días 16+)')
    
    def obtener_semana_mes(dia):
        if dia <= 7: return 'Sem 1 (Días 1-7)'
        elif dia <= 14: return 'Sem 2 (Días 8-14)'
        elif dia <= 21: return 'Sem 3 (Días 15-21)'
        elif dia <= 28: return 'Sem 4 (Días 22-28)'
        else: return 'Sem 5 (Días 29+)'
        
    df['Semana_del_Mes'] = df['Date'].dt.day.apply(obtener_semana_mes)

    mapa_dias = {0: '1-Lunes', 1: '2-Martes', 2: '3-Miércoles', 3: '4-Jueves', 4: '5-Viernes', 5: '6-Sábado', 6: '7-Domingo'}
    df['Dia_Semana'] = df['Date'].dt.dayofweek.map(mapa_dias)

    def es_dia_nomina(fecha):
        dia = fecha.day
        if dia in [14, 15, 16, 29, 30, 31, 1]: return "Dias_Nomina"
        else: return "Dias_Normales"
            
    df['Tipo_Dia'] = df['Date'].apply(es_dia_nomina)

    # --- AGRUPACIÓN Y VARIACIONES (SOBRE 2026) ---
    def obtener_totales_y_variaciones(df, columna_periodo):
        totales = df.groupby(['Holder', columna_periodo])['Abonos'].sum().unstack(fill_value=0)
        variacion = totales.diff(axis=1).fillna(0)
        return totales, variacion

    print("Calculando variaciones 2026...")
    abonos_mes, var_mes = obtener_totales_y_variaciones(df, 'Mes')
    abonos_sem, var_sem = obtener_totales_y_variaciones(df, 'Semana')
    abonos_dia, var_dia = obtener_totales_y_variaciones(df, 'Dia')

    def clasificar_tendencia(valor):
        if valor > 0: return 'Alza 📈'
        elif valor < 0: return 'Baja 📉'
        else: return 'Estable ➖'

    ultimo_mes = var_mes.columns[-1]
    resumen_mes = var_mes[[ultimo_mes]].copy()
    resumen_mes.columns = ['Variacion_Ultimo_Mes']
    resumen_mes['Tendencia_Mes'] = resumen_mes['Variacion_Ultimo_Mes'].apply(clasificar_tendencia)

    ultima_semana = var_sem.columns[-1]
    resumen_sem = var_sem[[ultima_semana]].copy()
    resumen_sem.columns = ['Variacion_Ultima_Semana']
    resumen_sem['Tendencia_Semana'] = resumen_sem['Variacion_Ultima_Semana'].apply(clasificar_tendencia)

    resumen_total = resumen_mes.join(resumen_sem).sort_values('Variacion_Ultimo_Mes', ascending=False)

    # --- PROYECCIÓN (SOBRE 2026) ---
    # Determinar mes actual dinámicamente
    max_date = df['Date'].max()
    current_year = max_date.year
    current_month = max_date.month

    print(f"Calculando proyecciones para el cierre de mes ({current_month}/{current_year})...")
    df_actual = df[(df['Date'].dt.year == current_year) & (df['Date'].dt.month == current_month)].copy()
    df_historico = df[~((df['Date'].dt.year == current_year) & (df['Date'].dt.month == current_month))]

    totales_por_mes = df_historico.groupby(['Holder', 'Mes'])['Abonos'].sum().reset_index()
    promedio_mensual = totales_por_mes.groupby('Holder')['Abonos'].mean().reset_index()
    promedio_mensual.rename(columns={'Abonos': 'Promedio_Mensual_Historico'}, inplace=True)

    totales_actual = df_actual.groupby('Holder')['Abonos'].sum().reset_index()
    totales_actual.rename(columns={'Abonos': 'Abonos_Actuales_Mes'}, inplace=True)

    dias_transcurridos = df_actual['Date'].dt.day.max()
    if pd.isna(dias_transcurridos): dias_transcurridos = 1
    dias_restantes = 31 - dias_transcurridos

    # Calcular totales por método de pago (Wuzi, BP, SPEI)
    totales_por_metodo = df.groupby('Holder')[['Wuzi', 'BP', 'SPEI']].sum().reset_index()

    lista_maestra = pd.DataFrame({'Holder': df['Holder'].unique()})
    proyeccion = pd.merge(lista_maestra, totales_actual, on='Holder', how='left').fillna(0)
    proyeccion = pd.merge(proyeccion, promedio_mensual, on='Holder', how='left').fillna(0)
    proyeccion = pd.merge(proyeccion, totales_por_metodo, on='Holder', how='left').fillna(0)

    proyeccion['Promedio_Diario'] = proyeccion['Abonos_Actuales_Mes'] / dias_transcurridos
    proyeccion['Proyeccion_Cierre_Mes'] = proyeccion['Abonos_Actuales_Mes'] + (proyeccion['Promedio_Diario'] * dias_restantes)
    proyeccion['Diferencia_vs_Promedio'] = proyeccion['Proyeccion_Cierre_Mes'] - proyeccion['Promedio_Mensual_Historico']

    def estatus_meta(row):
        if row['Promedio_Mensual_Historico'] == 0: return 'Nuevo Cliente (Sin historial previo)'
        elif row['Abonos_Actuales_Mes'] == 0: return '🚨 Inactivo en el mes actual (Riesgo de abandono)'
        elif row['Proyeccion_Cierre_Mes'] >= row['Promedio_Mensual_Historico']: return '✅ Llegará a la meta (Supera su promedio)'
        else: return '⚠️ Peligro (Por debajo de su promedio)'

    proyeccion['Diagnostico'] = proyeccion.apply(estatus_meta, axis=1)
    proyeccion = proyeccion.round(2).sort_values('Diferencia_vs_Promedio', ascending=False).set_index('Holder')

    # --- PATRONES (SOBRE 2026) ---
    print("Analizando patrones temporales y efecto de nómina...")
    patrones_quincena = df.groupby(['Holder', 'Quincena'])['Abonos'].sum().unstack(fill_value=0)
    if not patrones_quincena.empty: patrones_quincena['Quincena_Fuerte'] = patrones_quincena.idxmax(axis=1) 
    
    patrones_semana_mes = df.groupby(['Holder', 'Semana_del_Mes'])['Abonos'].sum().unstack(fill_value=0)
    if not patrones_semana_mes.empty: patrones_semana_mes['Semana_Fuerte'] = patrones_semana_mes.idxmax(axis=1) 

    patrones_dia_semana = df.groupby(['Holder', 'Dia_Semana'])['Abonos'].sum().unstack(fill_value=0)
    if not patrones_dia_semana.empty:
        columnas_ordenadas = [d for d in ['1-Lunes', '2-Martes', '3-Miércoles', '4-Jueves', '5-Viernes', '6-Sábado', '7-Domingo'] if d in patrones_dia_semana.columns]
        patrones_dia_semana = patrones_dia_semana[columnas_ordenadas]
        patrones_dia_semana['Dia_Fuerte'] = patrones_dia_semana.idxmax(axis=1)
        patrones_dia_semana['Dia_Fuerte'] = patrones_dia_semana['Dia_Fuerte'].str.split('-').str[1]

    holders_nomina = df.groupby(['Holder', 'Tipo_Dia']).agg(Ingresos=('Abonos', 'sum'), Dias=('Date', 'nunique')).reset_index()
    holders_nomina['Promedio_Diario'] = holders_nomina['Ingresos'] / holders_nomina['Dias']
    patron_nomina = holders_nomina.pivot(index='Holder', columns='Tipo_Dia', values='Promedio_Diario').fillna(0)
    
    if "Dias_Nomina" in patron_nomina.columns and "Dias_Normales" in patron_nomina.columns:
        patron_nomina['Diferencia_Absoluta'] = patron_nomina['Dias_Nomina'] - patron_nomina['Dias_Normales']
        patron_nomina = patron_nomina.sort_values('Diferencia_Absoluta', ascending=False).round(2)

    # =====================================================================================
    # --- NUEVO: SEMÁFORO EXCLUSIVO (BASADO SOLO EN LA HOJA 'SEMÁFORO') ---
    # =====================================================================================
    print("Calculando Semáforo de Salud utilizando EXCLUSIVAMENTE la hoja 'Semáforo'...")
    try:
        df_sem = pd.read_excel('TRX WU_BP.xlsx', sheet_name='Semáforo')
        cols = df_sem.columns.tolist()
        
        # Identificar columnas cronológicas y ordenarlas
        semanas_25 = [c for c in cols if str(c).startswith('Semana') and '2025' in str(c)]
        semanas_26 = [c for c in cols if str(c).startswith('Semana') and '2026' in str(c)]
        
        def obtener_num_semana(col_nombre):
            try: return int(col_nombre.split()[1])
            except: return 0

        semanas_25 = sorted(semanas_25, key=obtener_num_semana)
        semanas_26 = sorted(semanas_26, key=obtener_num_semana)
        timeline_cols = semanas_25 + semanas_26 
        
        df_sem_data = df_sem.iloc[1:].copy() # Saltar fila 0 (nombres de meses)
        col_holder = df_sem_data.columns[1]
        df_sem_data.rename(columns={col_holder: 'Holder'}, inplace=True)
        
        # Forzar a numérico y llenar huecos explícitamente con 0
        for col in timeline_cols:
            df_sem_data[col] = pd.to_numeric(df_sem_data[col], errors='coerce').fillna(0)
            
        estado_clientes = []
        for idx, row in df_sem_data.iterrows():
            holder = row['Holder']
            if pd.isna(holder) or str(holder).strip() == '' or str(holder).strip() == 'Total': 
                continue
            
            valores = row[timeline_cols].values
            indices_activos = np.where(valores > 0)[0]
            
            if len(indices_activos) == 0:
                estado = 'Sin Actividad'
                life, zero_w, tasa = 0, 0, 0
            else:
                first_week_idx = indices_activos[0]
                last_week_idx = len(valores) - 1 # Última semana registrada en las columnas
                
                # Semanas de vida: Desde el inicio hasta hoy
                life = last_week_idx - first_week_idx + 1
                
                # Cortar la historia exactamente al ciclo de vida del cliente
                historial_real = valores[first_week_idx:last_week_idx + 1]
                
                # Contar ceros dentro de su ciclo de vida
                zero_w = (historial_real == 0).sum()
                tasa = zero_w / life if life > 0 else 0
                
                # Evaluar las últimas 4 columnas (semanas) de toda la cuadrícula
                recent_activity = sum(valores[-4:]) > 0
                
                # Reglas de Negocio
                if life <= 4: estado = 'Nuevo 🌟'
                elif not recent_activity: estado = 'En Riesgo / Abandono 🚨'
                elif tasa > 0.25: estado = 'Intermitente ⚠️'
                else: estado = 'Constante ✅'
                    
            estado_clientes.append({
                'Holder': holder,
                'Semáforo': estado,
                'Semanas_De_Vida': life,
                'Semanas_En_Ceros': zero_w,
                'Porcentaje_Inactividad': f"{tasa:.1%}"
            })
            
        semaforo_df = pd.DataFrame(estado_clientes).sort_values(by='Semáforo', ascending=False).set_index('Holder')
    except Exception as e:
        print(f"Error al procesar hoja Semáforo: {e}")
        semaforo_df = pd.DataFrame()

    # --- PROCESAR HOJA DE RANKING (SI EXISTE) ---
    ranking_df = pd.DataFrame()
    try:
        xls = pd.ExcelFile('TRX WU_BP.xlsx')
        if 'Ranking' in xls.sheet_names:
            print("Cargando hoja de Ranking desde TRX WU_BP.xlsx...")
            ranking_df = pd.read_excel(xls, sheet_name='Ranking')
    except Exception as e:
        print(f"Nota: No se pudo leer la hoja 'Ranking': {e}")


    # --- GUARDAR TODO ---
    archivo_salida = 'Analisis_Historico_Variaciones.xlsx'
    print(f"Exportando resultados a {archivo_salida}...")
    with pd.ExcelWriter(archivo_salida) as writer:
        
        if not semaforo_df.empty:
            semaforo_df.to_excel(writer, sheet_name='Semaforo_Salud')
            
        if not ranking_df.empty:
            # Guardamos sin index para conservar las columnas limpias
            ranking_df.to_excel(writer, sheet_name='Ranking', index=False)
            
        proyeccion.to_excel(writer, sheet_name='Proyeccion_Cierre_Mes')
        resumen_total.to_excel(writer, sheet_name='Resumen_Tendencias_Actual')
        
        if 'patron_nomina' in locals() and not patron_nomina.empty:
            patron_nomina.to_excel(writer, sheet_name='Efecto_Nomina')
        patrones_quincena.to_excel(writer, sheet_name='Patrones_Quincena')
        patrones_semana_mes.to_excel(writer, sheet_name='Patrones_Semana_Mes')
        if not patrones_dia_semana.empty:
            patrones_dia_semana.to_excel(writer, sheet_name='Patrones_Dia_Semana')
            
        abonos_mes.to_excel(writer, sheet_name='Abonos_Mensuales')
        var_mes.to_excel(writer, sheet_name='Variacion_Mensual')
        abonos_sem.to_excel(writer, sheet_name='Abonos_Semanales')
        var_sem.to_excel(writer, sheet_name='Variacion_Semanal')
        abonos_dia.to_excel(writer, sheet_name='Abonos_Diarios')
        var_dia.to_excel(writer, sheet_name='Variacion_Diaria')
        
        # Guardar desglose semanal de métodos de pago
        wuzi_sem = df.groupby(['Holder', 'Semana'])['Wuzi'].sum().unstack(fill_value=0)
        bp_sem = df.groupby(['Holder', 'Semana'])['BP'].sum().unstack(fill_value=0)
        spei_sem = df.groupby(['Holder', 'Semana'])['SPEI'].sum().unstack(fill_value=0)
        
        wuzi_sem.to_excel(writer, sheet_name='Wuzi_Semanal')
        bp_sem.to_excel(writer, sheet_name='BP_Semanal')
        spei_sem.to_excel(writer, sheet_name='SPEI_Semanal')

if __name__ == '__main__':
    procesar_datos()