import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from data_processing import load_data, get_metrics_for_date, get_client_history, get_history_last_N_days
from export_utils import create_excel_report
import os

# Page config
st.set_page_config(page_title="Dashboard Alertas de Clientes", layout="wide")

# Custom CSS for premium look
st.markdown("""
<style>
    .kpi-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        text-align: center;
        margin-bottom: 20px;
        height: 140px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-title {
        color: #6c757d;
        font-size: 12px;
        font-weight: bold;
        text-transform: uppercase;
        line-height: 1.2;
    }
    .kpi-value {
        color: #212529;
        font-size: 28px;
        font-weight: bold;
        margin: 10px 0;
    }
    .kpi-red { border-left: 5px solid #dc3545; }
    .kpi-yellow { border-left: 5px solid #ffc107; }
    .kpi-green { border-left: 5px solid #28a745; }
    .kpi-blue { border-left: 5px solid #007bff; }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def get_raw_data_v2(file_mtime): 
    file_path = 'TRX WU_BP.xlsx'
    if not os.path.exists(file_path):
        return None, None
    df_raw, top_300 = load_data(file_path)
    return df_raw, top_300

st.title("🚨 Dashboard de Alertas Inteligentes de Clientes")

file_path = 'TRX WU_BP.xlsx'
file_mtime = os.path.getmtime(file_path) if os.path.exists(file_path) else 0

with st.spinner('Cargando datos crudos...'):
    df_raw, top_300 = get_raw_data_v2(file_mtime)

if df_raw is None:
    st.error("No se encontró el archivo 'TRX WU_BP.xlsx' en el directorio actual. Por favor, asegúrese de que el archivo esté presente.")
    st.stop()

# Filtros
st.sidebar.header("🎯 Filtros")
available_dates = sorted(df_raw['Date'].unique())
selected_date = st.sidebar.date_input("Fecha de Análisis", value=available_dates[-1].date(), min_value=available_dates[0].date(), max_value=available_dates[-1].date())
target_date = pd.to_datetime(selected_date)

# Calculate metrics ON THE FLY for the selected date
with st.spinner('Calculando métricas para la fecha seleccionada...'):
    df_latest = get_metrics_for_date(df_raw, top_300, target_date)
    
if df_latest.empty:
    st.warning("No hay datos para la fecha seleccionada o los datos son insuficientes.")
    st.stop()

# --- CALCULAR RANKING GLOBAL ---
# Se calcula antes de aplicar cualquier filtro para que el ranking nunca cambie
df_latest['Ranking_Global'] = df_latest['Acumulado_Historico'].rank(method='first', ascending=False).astype(int)

# Filtrar Top 300
show_only_top300 = st.sidebar.toggle("Mostrar solo Top 300", value=False)
if show_only_top300:
    df_latest = df_latest[df_latest['Es_Top_300'] == True]

# Filtro Riesgo
riesgos = ['Todos'] + list(df_latest['Nivel_riesgo'].unique())
selected_riesgo = st.sidebar.selectbox("Nivel de Riesgo", riesgos)
if selected_riesgo != 'Todos':
    df_latest = df_latest[df_latest['Nivel_riesgo'] == selected_riesgo]

# Filtro Alerta
alertas = ['Todas'] + list(df_latest['Alerta'].unique())
selected_alerta = st.sidebar.selectbox("Alerta", alertas)
if selected_alerta != 'Todas':
    df_latest = df_latest[df_latest['Alerta'] == selected_alerta]

# Crear contenedores para controlar el orden visual vs lógico
kpi_container = st.container()

st.markdown("---")
st.markdown("**Filtros de Equipo**")
filt_col1, filt_col2 = st.columns(2)

# Obtener listas únicas
directores = ['Todos'] + sorted([str(d) for d in df_latest.get('Director', []).unique() if pd.notna(d)])

with filt_col1:
    selected_director = st.selectbox("Director", directores)
    
if selected_director != 'Todos':
    vendedores = ['Todos'] + sorted([str(v) for v in df_latest[df_latest['Director'] == selected_director].get('Vendedor', []).unique() if pd.notna(v)])
else:
    vendedores = ['Todos'] + sorted([str(v) for v in df_latest.get('Vendedor', []).unique() if pd.notna(v)])
    
with filt_col2:
    selected_vendedor = st.selectbox("Vendedor", vendedores)

# Filtrar el dataframe lógicamente ANTES de renderizar los KPIs
if selected_director != 'Todos':
    df_latest = df_latest[df_latest['Director'] == selected_director]
if selected_vendedor != 'Todos':
    df_latest = df_latest[df_latest['Vendedor'] == selected_vendedor]

# Inicializar estado para el filtro de KPIs
if 'active_kpi_filter' not in st.session_state:
    st.session_state['active_kpi_filter'] = None
    
# CSS Hack para hacer las tarjetas enteras clickeables y ocultar el botón
st.markdown("""
<style>
/* Ocultar el botón y ponerlo encima de la tarjeta usando márgenes negativos en vez de absolute */
div.element-container:has(.kpi-overlay-marker) + div.element-container {
    margin-top: -110px !important;
    position: relative !important;
    z-index: 10 !important;
}
div.element-container:has(.kpi-overlay-marker) + div.element-container button {
    height: 110px !important;
    width: 100% !important;
    opacity: 0 !important;
    cursor: pointer !important;
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)
    
# Calcular KPIs con la data filtrada
with kpi_container:
    # --- RESUMEN GLOBAL ---
    st.header("🌎 Resumen Global")
    res_col1, res_col2 = st.columns(2)
    
    # Definición de mapeo de alertas con sus colores e iconos precisos (orden modificado)
    alert_map = {
        '1 día sin Trx': {'hex': '#ffcdd2', 'icon': '🔴'},
        '2 días seguidos sin Trx': {'hex': '#e57373', 'icon': '🔴'},
        '3 días seguidos sin Trx': {'hex': '#ef5350', 'icon': '🔴'},
        '4 a 7 días sin Trx': {'hex': '#f44336', 'icon': '🔴'},
        'Más de 7 días sin TRX': {'hex': '#c62828', 'icon': '🔴'},
        'Trx menor que el promedio diario': {'hex': '#ffb74d', 'icon': '🟠'},
        'Intermitente': {'hex': '#ffee58', 'icon': '🟡'},
        'Normal': {'hex': '#2196f3', 'icon': '🔵'},
        'Reactivado': {'hex': '#66bb6a', 'icon': '🟢'},
        'A la alza': {'hex': '#2e7d32', 'icon': '🟢'}
    }
    
    total_holders = len(df_latest)
    holders_con_trx = len(df_latest[df_latest['Abonos'] > 0]) if 'Abonos' in df_latest.columns else 0
    
    with res_col1:
        st.markdown(f'<div class="kpi-card kpi-blue" style="margin-bottom: 5px;"><div class="kpi-title">Total Holders</div><div class="kpi-value">{total_holders}</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="kpi-overlay-marker"></div>', unsafe_allow_html=True)
        if st.button("Todos", key="btn_todos", use_container_width=True):
            st.session_state['active_kpi_filter'] = None
            st.rerun()
            
    with res_col2:
        st.markdown(f'<div class="kpi-card" style="border-left: 5px solid #26a69a; margin-bottom: 5px;"><div class="kpi-title">Holders con TRX hoy</div><div class="kpi-value">{holders_con_trx}</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="kpi-overlay-marker"></div>', unsafe_allow_html=True)
        if st.button("Filtro TRX Hoy", key="btn_trx_hoy", use_container_width=True):
            st.session_state['active_kpi_filter'] = 'TRX_HOY'
            st.rerun()
            
    st.markdown("---")
    
    # --- KPIs POR TIPO DE ALERTA ---
    st.header("🚨 Alertas Activas")
    
    # Conteo de alertas
    alert_counts = df_latest['Alerta'].value_counts()
    
    # Crear filas de 5 columnas
    cols = st.columns(5)
    col_idx = 0
    
    for alerta, props in alert_map.items():
        count = alert_counts.get(alerta, 0)
        # Always show 0 counts for the very bad ones just to maintain visibility
        if count == 0 and alerta not in ['Más de 7 días sin TRX', '4 a 7 días sin Trx', '1 día sin Trx']: 
            continue 
            
        icon = props['icon']
        color_hex = props['hex']
            
        with cols[col_idx % 5]:
            # Tarjeta visual con color de borde izquierdo dinámico
            st.markdown(f'<div class="kpi-card" style="border-left: 5px solid {color_hex}; margin-bottom: 5px;"><div class="kpi-title">{icon} {alerta}</div><div class="kpi-value">{count}</div></div>', unsafe_allow_html=True)
            # Marcador oculto + Botón invisible
            st.markdown('<div class="kpi-overlay-marker"></div>', unsafe_allow_html=True)
            if st.button("Filtro", key=f"btn_{alerta}", use_container_width=True):
                st.session_state['active_kpi_filter'] = alerta
                st.rerun()
        col_idx += 1
        
        # Crear nueva fila si llegamos a 5
        if col_idx % 5 == 0 and col_idx < len(alert_map):
            cols = st.columns(5)

# Main layout
tab1, tab2 = st.tabs(["📋 Tabla Principal", "📊 Gráficos y Tendencias"])

with tab1:
    # Sort and format dataframe for display
    display_df = df_latest.copy()
    
    # Aplicar el filtro dinámico de las tarjetas KPI
    active_filter = st.session_state.get('active_kpi_filter')
    if active_filter is not None:
        if active_filter == 'FDS':
            display_df = display_df[display_df['Opera_Fin_De_Semana'] == True]
            st.info(f"Filtro de Tarjeta Activo: Mostrando exclusivamente a los **{len(display_df)}** clientes que operan en **Fines de Semana**. Haz clic en la tarjeta de **Total Holders** para quitar el filtro.")
        elif active_filter == 'TRX_HOY':
            display_df = display_df[display_df['Abonos'] > 0]
            st.info(f"Filtro de Tarjeta Activo: Mostrando exclusivamente a los **{len(display_df)}** clientes que tuvieron **TRX hoy**. Haz clic en la tarjeta de **Total Holders** para quitar el filtro.")
        else:
            display_df = display_df[display_df['Alerta'] == active_filter]
            st.info(f"Filtro de Tarjeta Activo: Mostrando exclusivamente a los **{len(display_df)}** clientes en **{active_filter}**. Haz clic en la tarjeta de **Total Holders** para quitar el filtro.")
    
    # Set the global ranking as the index
    display_df = display_df.set_index('Ranking_Global')
    display_df.index.name = 'Ranking'
    
    # We still sort by Score (risk) so the most at-risk clients appear at the top
    display_df = display_df.sort_values(by=['Score', 'Abonos'], ascending=[True, False])
    
    cols_to_show = ['Holder', 'Director', 'Vendedor', 'Es_Top_300', 'Opera_Fin_De_Semana', 'Abonos', 'Abonos_ayer', 'Promedio_diario_total', 'Variacion_promedio_pct', 'Dias_sin_transaccionar', 'Alerta']
    display_df = display_df[cols_to_show].rename(columns={
        'Abonos': 'Abonos Hoy',
        'Abonos_ayer': 'Abonos Ayer',
        'Promedio_diario_total': 'Promedio Diario Total',
        'Variacion_promedio_pct': 'Variación vs Promedio (%)',
        'Dias_sin_transaccionar': 'Días Inactivo',
        'Es_Top_300': 'Top 300',
        'Opera_Fin_De_Semana': 'Fin de Semana'
    })
    
    def style_dataframe(row):
        styles = [''] * len(row)
        alerta = row.get('Alerta', '')
        
        # Apply the exact hex background color with slightly lower opacity by hardcoding a lighter version if needed, or just use the hex directly if it works for text
        # But wait, these colors are intense! Let's keep them as background but maybe with white text if they are too dark.
        bg_color = alert_map.get(alerta, {}).get('hex', 'white')
        text_color = 'white' if alerta in ['Más de 7 días sin TRX', '4 a 7 días sin Trx', '3 días seguidos sin Trx', '2 días seguidos sin Trx', 'A la alza'] else 'black'
        
        # Color the 'Alerta' column background
        if 'Alerta' in row.index:
            idx_alerta = row.index.get_loc('Alerta')
            styles[idx_alerta] = f'background-color: {bg_color}; font-weight: 500; color: {text_color};'
            
        return styles
        
    st.dataframe(
        display_df.style.apply(style_dataframe, axis=1)
                        .format({
                            'Abonos Hoy': '${:,.2f}',
                            'Abonos Ayer': '${:,.2f}',
                            'Promedio Diario Total': '${:,.2f}',
                            'Variación vs Promedio (%)': '{:.2f}%'
                        }),

        use_container_width=True,
        height=600
    )

with tab2:
    st.subheader("Análisis Visual")
    
    c1, c2 = st.columns(2)
    with c1:
        # Pie chart -> nivel de riesgo
        if not df_latest.empty:
            fig_pie = px.pie(df_latest, names='Nivel_riesgo', title='Distribución por Nivel de Riesgo',
                             color='Nivel_riesgo', color_discrete_map={'Riesgo Alto': 'red', 'En Atención': 'orange', 'Saludable': 'green'})
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No hay datos para mostrar el gráfico circular.")
            
    with c2:
        # Bar chart -> Frecuencia de Alertas
        alert_counts = df_latest['Alerta'].value_counts().reset_index()
        alert_counts.columns = ['Alerta', 'Cantidad']
        
        if not alert_counts.empty:
            # Sort by count para que las barras más grandes queden arriba
            alert_counts = alert_counts.sort_values(by='Cantidad', ascending=True)
            
            # Usar alert_map definido arriba
            color_map = {k: v['hex'] for k, v in alert_map.items()}
            
            fig_bar = px.bar(
                alert_counts, 
                x='Cantidad', 
                y='Alerta', 
                orientation='h', 
                title='Distribución por Alerta',
                text_auto=True, 
                color='Alerta', 
                color_discrete_map=color_map
            )
            fig_bar.update_layout(showlegend=False, yaxis_title="")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No hay alertas para mostrar.")

    st.markdown("---")
    st.subheader("📉 Evolución Histórica de Score por Cliente")
    selected_client = st.selectbox("Seleccione un cliente para ver su evolución", df_latest['Holder'].unique())
    if selected_client:
        with st.spinner(f"Cargando historial de {selected_client}..."):
            client_history = get_client_history(df_raw, selected_client)
            
        if not client_history.empty:
            fig_line = px.line(client_history, x='Date', y='Score', title=f'Evolución del Score: {selected_client}', markers=True)
            fig_line.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="Límite Riesgo Alto")
            fig_line.add_hline(y=80, line_dash="dash", line_color="orange", annotation_text="Límite Atención")
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.warning("No hay historial suficiente para este cliente.")
        




# Exportar
st.sidebar.markdown("---")
st.sidebar.header("📤 Exportar Datos")

if st.sidebar.button("Generar Reporte Excel"):
    with st.spinner("Generando reporte (puede tardar un poco al calcular histórico de 30 días)..."):
        # Generar data para Excel
        df_kpi_summary = pd.DataFrame({
            'Métrica': ['Total Clientes', 'Riesgo Alto', 'En Atención', 'Saludables', 'Volumen en Riesgo'],
            'Valor': [len(df_latest), len(df_latest[df_latest['Nivel_riesgo'] == 'Riesgo Alto']), len(df_latest[df_latest['Nivel_riesgo'] == 'En Atención']), len(df_latest[df_latest['Nivel_riesgo'] == 'Saludable']), df_latest[df_latest['Nivel_riesgo'] == 'Riesgo Alto']['Promedio_diario_total'].sum() if 'Promedio_diario_total' in df_latest.columns else 0]
        })
        
        df_hist_30d = get_history_last_N_days(df_raw, top_300, target_date, days=30)
        
        st.session_state['excel_report_data'] = create_excel_report(df_latest, df_hist_30d, df_kpi_summary)
        st.session_state['excel_report_name'] = f"Reporte_Alertas_{selected_date.strftime('%Y%m%d')}.xlsx"
        st.rerun()

if 'excel_report_data' in st.session_state:
    st.sidebar.success("¡Reporte listo para descargar!")
    st.sidebar.download_button(
        label="✅ Descargar Reporte Generado",
        data=st.session_state['excel_report_data'],
        file_name=st.session_state['excel_report_name'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
