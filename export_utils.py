import pandas as pd
import io

def create_excel_report(df_latest, df_full, df_kpis):
    """
    Generates an Excel file in memory with multiple sheets.
    Returns bytes to be used in st.download_button.
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Alertas completas
        cols_main = ['Holder', 'Es_Top_300', 'Abonos', 'Abonos_ayer', 'Promedio_7d', 
                     'Variacion_promedio_pct', 'Dias_sin_transaccionar', 
                     'Tipo_alerta', 'Score', 'Nivel_riesgo']
        
        available_cols = [c for c in cols_main if c in df_latest.columns]
        
        df_alertas = df_latest[available_cols].rename(columns={
            'Abonos': 'Abonos_hoy',
            'Variacion_promedio_pct': 'Variacion_Promedio_%'
        })
        df_alertas.to_excel(writer, sheet_name='Alertas Completas', index=False)
        
        # Sheet 2: Top 300
        if 'Es_Top_300' in df_alertas.columns:
            df_top300 = df_alertas[df_alertas['Es_Top_300'] == True]
        else:
            df_top300 = df_alertas
        df_top300.to_excel(writer, sheet_name='Top 300', index=False)
        
        # Sheet 3: Resumen KPI
        df_kpis.to_excel(writer, sheet_name='Resumen KPI', index=False)
        
        # Sheet 4: Histórico de score
        cols_hist = ['Holder', 'Date', 'Score', 'Abonos', 'Tipo_alerta']
        available_hist_cols = [c for c in cols_hist if c in df_full.columns]
        df_historico = df_full[available_hist_cols]
        
        # Excel limit is 1,048,576 rows per sheet. 
        # If we exceed it, we filter to recent dates or simply truncate.
        MAX_ROWS = 1000000
        if len(df_historico) > MAX_ROWS:
            # Sort by Date descending to keep the most recent ones
            df_historico = df_historico.sort_values('Date', ascending=False).head(MAX_ROWS)
            
        df_historico.to_excel(writer, sheet_name='Histórico Score', index=False)
        
    processed_data = output.getvalue()
    return processed_data
