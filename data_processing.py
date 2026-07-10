import pandas as pd
import numpy as np

def load_data(filepath):
    print("Cargando datos raw...")
    df = pd.read_excel(filepath)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
    
    # Filtrar fechas basura (ej. 1970-01-01 por culpa de un 0 en el excel) o nulas
    df = df.dropna(subset=['Date'])
    df = df[df['Date'] > pd.Timestamp('2000-01-01')]
    
    if 'Holder' not in df.columns or 'Abonos' not in df.columns:
        raise ValueError("El archivo debe contener las columnas 'Date', 'Holder', y 'Abonos'.")
        
    df = df[df['Holder'].astype(str).str.lower() != 'total']
    df['Abonos'] = df['Abonos'].fillna(0)
    
    df = df.groupby(['Date', 'Holder'], as_index=False)['Abonos'].sum()
    
    total_acumulado = df.groupby('Holder')['Abonos'].sum()
    top_300 = set(total_acumulado.nlargest(300).index)
    
    return df, top_300

def get_metrics_for_date(df_raw, top_300, target_date):
    """
    Calcula los scores "On-the-fly" solo para el target_date.
    """
    d_0 = pd.to_datetime(target_date)
    d_1 = d_0 - pd.Timedelta(days=1)
    d_2 = d_0 - pd.Timedelta(days=2)
    d_6 = d_0 - pd.Timedelta(days=6)
    d_7 = d_0 - pd.Timedelta(days=7)
    
    df_past = df_raw[df_raw['Date'] <= d_0]
    
    # Solo analizar clientes que hayan tenido al menos 1 transacción histórica (>0) hasta la fecha seleccionada.
    total_hasta_fecha = df_past.groupby('Holder')['Abonos'].sum()
    all_holders = total_hasta_fecha[total_hasta_fecha > 0].index
    
    if len(all_holders) == 0:
        return pd.DataFrame()
    
    # --- Base calculations ---
    # Hoy
    abonos_hoy = df_past[df_past['Date'] == d_0].groupby('Holder')['Abonos'].sum()
    abonos_hoy = abonos_hoy.reindex(all_holders).fillna(0)
    
    # Ayer
    abonos_ayer = df_past[df_past['Date'] == d_1].groupby('Holder')['Abonos'].sum()
    abonos_ayer = abonos_ayer.reindex(all_holders).fillna(0)
    
    # Promedio Diario Total
    first_active = df_past[df_past['Abonos'] > 0].groupby('Holder')['Date'].min()
    first_active = first_active.reindex(all_holders)
    dias_totales_vida = (d_0 - first_active).dt.days.clip(lower=0) + 1
    promedio_diario_total = (total_hasta_fecha.reindex(all_holders) / dias_totales_vida).fillna(0)
    
    # Días sin transaccionar
    last_active = df_past[df_past['Abonos'] > 0].groupby('Holder')['Date'].max()
    last_active = last_active.reindex(all_holders)
    dias_inactivo = (d_0 - last_active).dt.days.fillna(999)
    
    # Tendencia últimos 3 días (d_0, d_1, d_2)
    abonos_d2 = df_past[df_past['Date'] == d_2].groupby('Holder')['Abonos'].sum().reindex(all_holders).fillna(0)
    tendencia_3d = (abonos_hoy < abonos_ayer) & (abonos_ayer < abonos_d2) & (abonos_ayer > 0)
    
    # Intermitencia (días en $0 en los últimos 7 días)
    abonos_7d_df = df_past[df_past['Date'] >= d_7].groupby(['Holder', 'Date'])['Abonos'].sum().unstack(fill_value=0)
    abonos_7d_df = abonos_7d_df.reindex(all_holders, fill_value=0)
    dias_en_cero = (abonos_7d_df == 0).sum(axis=1)
    es_intermitente = (dias_en_cero >= 2) & (dias_en_cero <= 5)
    
    # Estaba inactivo ayer (para Reactivados)
    last_active_before_today = df_past[(df_past['Date'] < d_0) & (df_past['Abonos'] > 0)].groupby('Holder')['Date'].max()
    dias_inactivo_ayer = (d_1 - last_active_before_today.reindex(all_holders)).dt.days.fillna(999)
    estaba_inactivo_ayer = dias_inactivo_ayer >= 3

    # Opera Fin de Semana
    weekend_txs = df_past[(df_past['Date'].dt.dayofweek >= 5) & (df_past['Abonos'] > 0)].groupby('Holder').size()
    opera_fin_de_semana = weekend_txs.reindex(all_holders).fillna(0) > 0

    res = pd.DataFrame(index=all_holders)
    res.index.name = 'Holder'
    res['Date'] = d_0
    res['Abonos'] = abonos_hoy
    res['Abonos_ayer'] = abonos_ayer
    res['Promedio_diario_total'] = promedio_diario_total
    res['Es_Top_300'] = res.index.isin(top_300)
    res['Variacion_ayer_pct'] = np.where(res['Abonos_ayer'] == 0, 0, (res['Abonos'] - res['Abonos_ayer']) / res['Abonos_ayer'] * 100)
    res['Variacion_promedio_pct'] = np.where(res['Promedio_diario_total'] == 0, 0, (res['Abonos'] - res['Promedio_diario_total']) / res['Promedio_diario_total'] * 100)
    res['Dias_sin_transaccionar'] = dias_inactivo
    res['Tendencia_negativa_3d'] = tendencia_3d
    res['Es_Intermitente'] = es_intermitente
    res['Estaba_inactivo_ayer'] = estaba_inactivo_ayer
    res['Opera_Fin_De_Semana'] = opera_fin_de_semana
    res['Acumulado_Historico'] = total_hasta_fecha.reindex(all_holders).fillna(0)
    
    res = res.reset_index()
    
    # Intentar cargar Ranking.csv para Director y Vendedor
    try:
        import os
        ranking_path = os.path.join('webapp', 'Reportes_Individuales_CSV', 'Ranking.csv')
        
        if os.path.exists(ranking_path):
            df_ranking = pd.read_csv(ranking_path, encoding='utf-8')
            if 'Holder' in df_ranking.columns and 'Director' in df_ranking.columns and 'Vendedor' in df_ranking.columns:
                # Normalizar columnas para el merge (mayúsculas y sin espacios a los lados)
                res['_merge_key'] = res['Holder'].astype(str).str.strip().str.upper()
                df_ranking['_merge_key'] = df_ranking['Holder'].astype(str).str.strip().str.upper()
                
                # Quitar duplicados en _merge_key en ranking para evitar cruces dobles
                df_ranking_dedup = df_ranking[['_merge_key', 'Director', 'Vendedor']].drop_duplicates(subset=['_merge_key'])
                
                # Merge con res usando la clave normalizada
                res = pd.merge(res, df_ranking_dedup, on='_merge_key', how='left')
                res = res.drop(columns=['_merge_key'])
                
                res['Director'] = res['Director'].fillna('Sin Director')
                res['Vendedor'] = res['Vendedor'].fillna('Sin Vendedor')
            else:
                res['Director'] = 'Desconocido'
                res['Vendedor'] = 'Desconocido'
        else:
            res['Director'] = 'Desconocido'
            res['Vendedor'] = 'Desconocido'
    except Exception as e:
        res['Director'] = 'Error'
        res['Vendedor'] = 'Error'
        
    res = generate_alerts_and_score(res)
    return res

def generate_alerts_and_score(df_full):
    df_full['Score'] = 100
    df_full['Alerta'] = 'Normal'
    df_full['Nivel_riesgo'] = 'Saludable'
    
    m_inactivo_1 = df_full['Dias_sin_transaccionar'] == 1
    m_inactivo_2 = df_full['Dias_sin_transaccionar'] == 2
    m_inactivo_3 = df_full['Dias_sin_transaccionar'] == 3
    m_inactivo_4_7 = (df_full['Dias_sin_transaccionar'] >= 4) & (df_full['Dias_sin_transaccionar'] <= 7)
    m_inactivo_mas_7 = df_full['Dias_sin_transaccionar'] > 7
    
    m_caida = (df_full['Variacion_promedio_pct'] <= -30) & (df_full['Abonos'] > 0)
    m_tendencia = df_full['Tendencia_negativa_3d'] & (df_full['Abonos'] > 0)
    m_reactivado = (df_full['Abonos'] > 0) & df_full['Estaba_inactivo_ayer']
    m_crecimiento = (df_full['Variacion_promedio_pct'] >= 30) & (df_full['Abonos'] > 0) & ~m_reactivado
    m_intermitente = df_full['Es_Intermitente'] & (df_full['Abonos'] > 0)
    
    df_full.loc[m_inactivo_1, 'Score'] -= 25
    df_full.loc[m_inactivo_2, 'Score'] -= 40
    df_full.loc[m_inactivo_3, 'Score'] -= 50
    df_full.loc[m_inactivo_4_7, 'Score'] -= 60
    df_full.loc[m_inactivo_mas_7, 'Score'] -= 85
    
    df_full.loc[m_caida, 'Score'] -= 25
    df_full.loc[m_tendencia, 'Score'] -= 15
    df_full.loc[m_intermitente, 'Score'] -= 10
    df_full.loc[m_crecimiento, 'Score'] += 10
    df_full.loc[m_reactivado, 'Score'] += 15 
    
    df_full['Score'] = df_full['Score'].clip(lower=0, upper=100)
    
    df_full.loc[m_intermitente, 'Alerta'] = 'Intermitente'
    df_full.loc[m_tendencia, 'Alerta'] = 'Trx menor que el promedio diario'
    df_full.loc[m_caida, 'Alerta'] = 'Trx menor que el promedio diario'
    
    df_full.loc[m_inactivo_1, 'Alerta'] = '1 día sin Trx'
    df_full.loc[m_inactivo_2, 'Alerta'] = '2 días seguidos sin Trx'
    df_full.loc[m_inactivo_3, 'Alerta'] = '3 días seguidos sin Trx'
    df_full.loc[m_inactivo_4_7, 'Alerta'] = '4 a 7 días sin Trx'
    df_full.loc[m_inactivo_mas_7, 'Alerta'] = 'Más de 7 días sin TRX'
    
    df_full.loc[m_crecimiento, 'Alerta'] = 'A la alza'
    df_full.loc[m_reactivado, 'Alerta'] = 'Reactivado'
    
    # 80 is now 'En Atención'
    df_full.loc[df_full['Score'] <= 80, 'Nivel_riesgo'] = 'En Atención'
    df_full.loc[df_full['Score'] < 50, 'Nivel_riesgo'] = 'Riesgo Alto'
    
    return df_full

def get_client_history(df_raw, client):
    """
    Calculates historical scores only for a single client (very fast).
    """
    client_df = df_raw[df_raw['Holder'] == client].copy()
    if client_df.empty:
        return pd.DataFrame()
        
    dates = pd.date_range(start=client_df['Date'].min(), end=client_df['Date'].max(), freq='D', name='Date')
    pivot_df = client_df.pivot(index='Date', columns='Holder', values='Abonos').reindex(dates).fillna(0)
    
    abonos_ayer = pivot_df.shift(1).fillna(0)
    promedio_7d = pivot_df.rolling(7, min_periods=1).mean()
    abonos_d_minus_2 = pivot_df.shift(2).fillna(0)
    
    transicion_a_cero = ((pivot_df == 0) & (abonos_ayer > 0)).astype(int)
    transiciones_7d = transicion_a_cero.rolling(7, min_periods=1).sum().fillna(0)
    
    s_abonos = pivot_df.unstack()
    df_hist = s_abonos.reset_index(name='Abonos')
    df_hist.columns = ['Holder', 'Date', 'Abonos']
    
    df_hist['Abonos_ayer'] = abonos_ayer.unstack().values
    df_hist['Promedio_7d'] = promedio_7d.unstack().values
    df_hist['Abonos_D_minus_2'] = abonos_d_minus_2.unstack().values
    df_hist['Transiciones_7d'] = transiciones_7d.unstack().values
    
    df_hist['Variacion_ayer_pct'] = np.where(df_hist['Abonos_ayer'] == 0, 0, 
                                            (df_hist['Abonos'] - df_hist['Abonos_ayer']) / df_hist['Abonos_ayer'] * 100)
    df_hist['Variacion_promedio_pct'] = np.where(df_hist['Promedio_7d'] == 0, 0, 
                                                (df_hist['Abonos'] - df_hist['Promedio_7d']) / df_hist['Promedio_7d'] * 100)
                                                
    df_hist['Tendencia_negativa_3d'] = (df_hist['Abonos_D_minus_2'] > df_hist['Abonos_ayer']) & \
                                       (df_hist['Abonos_ayer'] > df_hist['Abonos']) & \
                                       (df_hist['Abonos_ayer'] > 0)
                                       
    df_hist['Es_Intermitente'] = df_hist['Transiciones_7d'] >= 2
    
    is_zero = (df_hist['Abonos'] == 0).astype(int)
    block = (is_zero == 0).cumsum()
    df_hist['Dias_sin_transaccionar'] = is_zero.groupby(block).cumsum()
    df_hist['Estaba_inactivo_ayer'] = df_hist['Dias_sin_transaccionar'].shift(1).fillna(0) >= 3
    
    df_hist = generate_alerts_and_score(df_hist)
    return df_hist

def get_history_last_N_days(df_raw, top_300, target_date, days=30):
    """
    Calculates metrics for the last N days for all holders to export.
    """
    dfs = []
    d_end = pd.to_datetime(target_date)
    for i in range(days):
        d_curr = d_end - pd.Timedelta(days=i)
        if d_curr >= df_raw['Date'].min():
            daily_metrics = get_metrics_for_date(df_raw, top_300, d_curr)
            dfs.append(daily_metrics)
    
    if len(dfs) == 0:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)
