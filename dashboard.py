import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go

# Configuración de la página
st.set_page_config(page_title="Dashboard de Inteligencia de Clientes", layout="wide")

# Estilos CSS personalizados para tarjetas de KPI premium
st.markdown("""
<style>
    .kpi-card {
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        text-align: center;
        margin-bottom: 20px;
        min-height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .kpi-title {
        color: #595959;
        font-size: 11px;
        font-weight: bold;
        text-transform: uppercase;
        margin-bottom: 5px;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: bold;
        color: #1f4e78;
        margin-bottom: 2px;
        line-height: 1;
    }
    .kpi-green { border-left: 5px solid #2e7d32; background-color: #e8f5e9; }
    .kpi-red { border-left: 5px solid #c65911; background-color: #fbe9e7; }
    .kpi-blue { border-left: 5px solid #1f4e78; background-color: #e3f2fd; }
    .kpi-grey { border-left: 5px solid #595959; background-color: #f5f5f5; }
</style>
""", unsafe_allow_html=True)

# ----------------- CARGA DE DATOS -----------------

# 1. Cargar datos operativos (Pestaña 1 - Detalle Diario 2026)
@st.cache_data
def cargar_datos_principales():
    df = pd.read_excel('TRX WU_BP.xlsx', sheet_name=0)
    df = df[df['Date'] != 'Total']
    df = df[df['Holder'] != 'Total']
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    
    # Crear periodos operativos
    df['Mes'] = df['Date'].dt.to_period('M').astype(str)
    df['Semana'] = df['Date'].dt.to_period('W').astype(str)
    df['Quincena'] = df['Date'].dt.day.apply(lambda x: '1ra Quincena (1-15)' if x <= 15 else '2da Quincena (16+)')
    
    def semana_mes(dia):
        if dia <= 7: return 'Sem 1'
        elif dia <= 14: return 'Sem 2'
        elif dia <= 21: return 'Sem 3'
        elif dia <= 28: return 'Sem 4'
        else: return 'Sem 5'
    df['Semana_del_Mes'] = df['Date'].dt.day.apply(semana_mes)
    return df

# 2. Cargar datos históricos de la hoja Semáforo (Semanas de 2025 y 2026)
@st.cache_data
def cargar_datos_semaforo():
    df_sem = pd.read_excel('TRX WU_BP.xlsx', sheet_name='Semáforo')
    
    cols = df_sem.columns.tolist()
    # Filtrar columnas que corresponden a las semanas estructuradas
    semanas_25 = [c for c in cols if str(c).startswith('Semana') and '2025' in str(c)]
    semanas_26 = [c for c in cols if str(c).startswith('Semana') and '2026' in str(c)]
    
    # La fila 0 contiene las etiquetas de los meses correspondientes
    meses_row = df_sem.iloc[0]
    mapping = []
    for c in semanas_25 + semanas_26:
        mes = str(meses_row[c]).strip() if pd.notna(meses_row[c]) else "Desconocido"
        year = '2025' if '2025' in c else '2026'
        
        # Extraer el número de semana de forma segura
        partes_columna = c.split()
        week_num = int(partes_columna[1]) if len(partes_columna) >= 2 and partes_columna[1].isdigit() else 0
        
        mapping.append({'Col': c, 'Mes': mes, 'Año': year, 'Semana': week_num})
    
    df_map = pd.DataFrame(mapping)
    
    # El set de datos real comienza después de la fila de meses
    df_data = df_sem.iloc[1:].copy()
    col_holder = df_data.columns[1]  # La segunda columna mapea el Holder
    df_data.rename(columns={col_holder: 'Holder'}, inplace=True)
    
    # Transformar a valores numéricos limpios
    for col in df_map['Col']:
        df_data[col] = pd.to_numeric(df_data[col], errors='coerce').fillna(0)
        
    return df_data, df_map

# 3. Cargar Reporte de Salud (Reporte_Salud_Clientes_2026.xlsx)
@st.cache_data
def cargar_reporte_salud():
    salud_file = 'Reporte_Salud_Clientes_2026.xlsx'
    if not os.path.exists(salud_file):
        return None, None
    
    # Leer resumen
    df_res = pd.read_excel(salud_file, sheet_name='Resumen de Salud', header=3)
    df_res_clean = df_res.iloc[0:4, [6, 7, 8]].copy()
    df_res_clean.columns = ['Estado de Salud', 'Clientes', '% Participación']
    
    # Leer pestañas de detalle
    detalles = {}
    for name in ['Ha Subido', 'Ha Bajado', 'Constante', 'Perdidos']:
        detalles[name] = pd.read_excel(salud_file, sheet_name=name, header=3)
        
    return df_res_clean, detalles

# 4. Cargar Reporte de Perdidos (Reporte_Clientes_Perdidos.xlsx)
@st.cache_data
def cargar_reporte_perdidos():
    perdidos_file = 'Reporte_Clientes_Perdidos.xlsx'
    if not os.path.exists(perdidos_file):
        return None
    
    detalles = {}
    resumenes = {}
    for name in ['Perdidos 2025', 'Activos 2025 - Perdidos 2026', 'Activos 2026 - Perdidos 2026']:
        detalles[name] = pd.read_excel(perdidos_file, sheet_name=name, header=3)
        
        # Leer la tabla de resumen mensual de cada pestaña (columnas F-K, header en fila 4)
        df_res = pd.read_excel(perdidos_file, sheet_name=name, header=3)
        df_res_clean = df_res.iloc[:, [5, 6, 7, 8, 9, 10]].dropna(subset=[df_res.columns[5]])
        df_res_clean = df_res_clean[df_res_clean.iloc[:, 0].astype(str).str.lower() != 'total únicos']
        df_res_clean.columns = ["Mes", "Total Clientes", "Clientes Activos", "% Activos", "Clientes Perdidos", "% Perdidos"]
        resumenes[name] = df_res_clean
        
    return detalles, resumenes

# --- EJECUTAR PROCESOS DE LECTURA ---
df_main = cargar_datos_principales()

# ----------------- NAVEGACIÓN LATERAL -----------------
st.sidebar.title("📌 Menú de Navegación")
vista_seleccionada = st.sidebar.radio(
    "Selecciona un Reporte:",
    [
        "📈 Análisis de Transacciones y Proyecciones",
        "💚 Salud de la Cartera (Health Score)",
        "🚨 Clientes Perdidos (Churn)"
    ]
)

# ----------------- VISTA 1: TRANSACCIONES Y PROYECCIONES -----------------
if vista_seleccionada == "📈 Análisis de Transacciones y Proyecciones":
    st.title("📊 Inteligencia de Negocio: Análisis de Clientes")
    st.markdown("Web App optimizada para cruce de datos diarios (2026) e históricos (Semáforo 2025-2026).")
    
    try:
        df_sem, df_map = cargar_datos_semaforo()
        tiene_semaforo = True
    except Exception as e:
        st.error(f"Error al procesar la estructura de la hoja 'Semáforo': {e}")
        tiene_semaforo = False

    st.sidebar.divider()
    st.sidebar.subheader("⚙️ Filtro de Cuentas")
    lista_clientes = ['Todos (Vista Global)', 'Top 300 (Vista Global)'] + sorted(df_main['Holder'].dropna().astype(str).unique().tolist())
    cliente_seleccionado = st.sidebar.selectbox("Selecciona un Cliente para analizar:", lista_clientes)
    
    # Filtrado dinámico de datos de acuerdo a la selección
    if cliente_seleccionado == 'Todos (Vista Global)':
        df_mostrar = df_main
        if tiene_semaforo: df_sem_mostrar = df_sem
        st.subheader("Análisis Global (Todo el Portafolio)")
        
    elif cliente_seleccionado == 'Top 300 (Vista Global)':
        top_300_clientes = df_main.groupby('Holder')['Abonos'].sum().nlargest(300).index
        df_mostrar = df_main[df_main['Holder'].isin(top_300_clientes)]
        if tiene_semaforo: df_sem_mostrar = df_sem[df_sem['Holder'].isin(top_300_clientes)]
        st.subheader("Análisis Global (Top 300 Clientes VIP)")
        
    else:
        df_mostrar = df_main[df_main['Holder'] == cliente_seleccionado]
        if tiene_semaforo: df_sem_mostrar = df_sem[df_sem['Holder'] == cliente_seleccionado]
        st.subheader(f"Análisis Individual: {cliente_seleccionado}")
    
    st.divider()
    
    # Gráficas Comparativas YoY
    if tiene_semaforo:
        st.subheader("⚖️ Comparativa Interanual Histórica (2025 vs 2026)")
        
        sums = df_sem_mostrar[df_map['Col']].sum()
        df_plot = pd.DataFrame({'Col': sums.index, 'Total': sums.values})
        df_plot = pd.merge(df_plot, df_map, on='Col')
        
        col_yoy1, col_yoy2 = st.columns(2)
        
        with col_yoy1:
            st.markdown("**Comparativa Mensual Agrupada**")
            monthly_comp = df_plot.groupby(['Mes', 'Año'])['Total'].sum().unstack(fill_value=0)
            
            meses_orden = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            meses_presentes = [m for m in meses_orden if m in monthly_comp.index]
            monthly_comp = monthly_comp.loc[meses_presentes]
            
            st.bar_chart(monthly_comp, color=["#1E88E5", "#FFC107"][:len(monthly_comp.columns)])
            
        with col_yoy2:
            st.markdown("**Comparativa Semanal de Líneas (Semana 1 a 53)**")
            weekly_comp = df_plot.groupby(['Semana', 'Año'])['Total'].sum().unstack(fill_value=0)
            st.line_chart(weekly_comp, color=["#1E88E5", "#FFC107"][:len(weekly_comp.columns)])
            
        st.divider()
    
    # Métricas Operativas y Run Rate
    st.subheader("🎯 Proyección y Ritmo Actual (Mayo 2026)")
    col1, col2, col3 = st.columns(3)
    
    df_mayo = df_mostrar[(df_mostrar['Date'].dt.year == 2026) & (df_mostrar['Date'].dt.month == 5)]
    df_historico = df_mostrar[~((df_mostrar['Date'].dt.year == 2026) & (df_mostrar['Date'].dt.month == 5))]
    
    abonos_mayo = df_mayo['Abonos'].sum()
    dias_transcurridos = df_mayo['Date'].dt.day.max() if not df_mayo.empty else 1
    if pd.isna(dias_transcurridos): dias_transcurridos = 1
    proyeccion = abonos_mayo + ((abonos_mayo / dias_transcurridos) * (31 - dias_transcurridos))
    promedio_mensual = df_historico.groupby('Mes')['Abonos'].sum().mean() if not df_historico.empty else 0
    
    col1.metric("Ingresos Actuales (Mayo)", f"${abonos_mayo:,.2f}")
    col2.metric("Proyección al 31 de Mayo", f"${proyeccion:,.2f}", f"{proyeccion - promedio_mensual:,.2f} vs Promedio 2026")
    col3.metric("Promedio Mensual (2026)", f"${promedio_mensual:,.2f}")
    
    st.divider()
    
    # Patrones de Liquidez
    st.subheader("🔍 Patrones de Comportamiento Intra-mes (2026)")
    col_patron1, col_patron2 = st.columns(2)
    
    with col_patron1:
        st.markdown("**Distribución por Quincena**")
        patron_quin = df_mostrar.groupby('Quincena')['Abonos'].sum().reset_index()
        st.bar_chart(patron_quin.set_index('Quincena'), color="#4a4e4d")
        
    with col_patron2:
        st.markdown("**Comportamiento por Semana del Mes**")
        patron_sem = df_mostrar.groupby('Semana_del_Mes')['Abonos'].sum().reset_index()
        st.bar_chart(patron_sem.set_index('Semana_del_Mes'), color="#f2a900")

# ----------------- VISTA 2: HEALTH SCORE -----------------
elif vista_seleccionada == "💚 Salud de la Cartera (Health Score)":
    st.title("💚 Dashboard de Salud de la Cartera 2026")
    st.markdown("Análisis interactivo de salud de clientes activos basado en variaciones contra histórico.")
    
    df_res_health, detalles_health = cargar_reporte_salud()
    
    if df_res_health is None:
        st.error("No se encontró el reporte `Reporte_Salud_Clientes_2026.xlsx`. Corre `ejecutar_analisis.bat` primero para generarlo.")
    else:
        # Calcular métricas dinámicas
        num_creciendo = len(detalles_health['Ha Subido'])
        num_caida = len(detalles_health['Ha Bajado'])
        num_estables = len(detalles_health['Constante'])
        num_bajas = len(detalles_health['Perdidos'])
        total_clientes = num_creciendo + num_caida + num_estables + num_bajas
        sana_clientes = num_creciendo + num_estables
        
        # Tarjetas de KPI
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown(f'<div class="kpi-card kpi-green"><div class="kpi-title">CRECIENDO 📈</div><div class="kpi-value">{num_creciendo}</div><div style="font-size:13px; font-weight:bold; color:#2e7d32;">{num_creciendo/total_clientes*100:.1f}%</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="kpi-card kpi-red"><div class="kpi-title">EN CAÍDA 📉</div><div class="kpi-value">{num_caida}</div><div style="font-size:13px; font-weight:bold; color:#c65911;">{num_caida/total_clientes*100:.1f}%</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="kpi-card kpi-blue"><div class="kpi-title">ESTABLES ➖</div><div class="kpi-value">{num_estables}</div><div style="font-size:13px; font-weight:bold; color:#1f4e78;">{num_estables/total_clientes*100:.1f}%</div></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="kpi-card kpi-grey"><div class="kpi-title">BAJAS 🚨</div><div class="kpi-value">{num_bajas}</div><div style="font-size:13px; font-weight:bold; color:#595959;">{num_bajas/total_clientes*100:.1f}%</div></div>', unsafe_allow_html=True)
        with col5:
            st.markdown(f'<div class="kpi-card" style="border-left: 5px solid #1b5e20; background-color: #C8E6C9;"><div class="kpi-title" style="color:#1b5e20; font-weight:bold;">CARTERA SANA 💚</div><div class="kpi-value" style="color:#1b5e20;">{sana_clientes}</div><div style="font-size:13px; font-weight:bold; color:#1b5e20;">{sana_clientes/total_clientes*100:.1f}%</div></div>', unsafe_allow_html=True)
            
        st.divider()
        
        # Tabla y Gráfico
        col_t1, col_t2 = st.columns([2, 3])
        
        with col_t1:
            st.subheader("Tabla Resumen de Salud")
            st.dataframe(
                df_res_health.style.format({'Clientes': '{:,}', '% Participación': '{:.1%}'}),
                use_container_width=True,
                hide_index=True
            )
            st.info(f"💡 **Tasa de Salud**: El **{sana_clientes/total_clientes*100:.1f}%** de la cartera está sana (estable o en crecimiento).")
            
        with col_t2:
            fig_doughnut = px.pie(
                df_res_health,
                names='Estado de Salud',
                values='Clientes',
                color='Estado de Salud',
                color_discrete_map={
                    'Ha subido': '#A9DFBF',
                    'Ha bajado': '#FADBD8',
                    'Constante': '#AED6F1',
                    'Perdido': '#E5E7E9'
                },
                hole=0.45
            )
            fig_doughnut.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
            st.plotly_chart(fig_doughnut, use_container_width=True)
            
        st.divider()
        
        # Detalle de Clientes
        st.subheader("🔍 Detalle de Clientes por Estado de Salud")
        estado_seleccionado = st.selectbox(
            "Selecciona un Estado para ver sus holders:",
            ["Ha Subido", "Ha Bajado", "Constante", "Perdidos"]
        )
        
        df_detalles = detalles_health[estado_seleccionado].copy()
        
        # Buscador de texto
        busqueda = st.text_input(f"Buscar Holder en '{estado_seleccionado}':")
        if busqueda:
            df_detalles = df_detalles[df_detalles['Nombre del Holder'].astype(str).str.contains(busqueda, case=False, na=False)]
            
        st.markdown(f"Mostrando **{len(df_detalles)}** de **{len(detalles_health[estado_seleccionado])}** clientes.")
        
        if estado_seleccionado in ["Ha Subido", "Ha Bajado", "Constante"]:
            st.dataframe(
                df_detalles[['Nombre del Holder', 'Director', 'Promedio Inicial (W1-21)', 'Promedio Reciente (W22-25)', 'Variación (%)']].style.format({
                    'Promedio Inicial (W1-21)': '${:,.2f}',
                    'Promedio Reciente (W22-25)': '${:,.2f}',
                    'Variación (%)': '{:+.1%}'
                }),
                use_container_width=True,
                hide_index=True
            )
        else: # Perdidos
            st.dataframe(
                df_detalles[['Nombre del Holder', 'Director', 'Última Semana Activo', 'Último Mes Activo', 'Promedio Semanal Histórico']].style.format({
                    'Promedio Semanal Histórico': '${:,.2f}'
                }),
                use_container_width=True,
                hide_index=True
            )

# ----------------- VISTA 3: CLIENTES PERDIDOS -----------------
elif vista_seleccionada == "🚨 Clientes Perdidos (Churn)":
    st.title("🚨 Dashboard de Clientes Perdidos")
    st.markdown("Seguimiento y análisis mensual de cuentas inactivas durante las últimas semanas o de periodos anteriores.")
    
    datos_perdidos = cargar_reporte_perdidos()
    
    if datos_perdidos is None:
        st.error("No se encontró el reporte `Reporte_Clientes_Perdidos.xlsx`. Corre `ejecutar_analisis.bat` primero para generarlo.")
    else:
        detalles_perd, resumenes_perd = datos_perdidos
        
        # Selector de Categoría
        categoria_sel = st.selectbox(
            "Selecciona una Categoría de Pérdida:",
            [
                "Activos 2025 - Perdidos 2026",
                "Activos 2026 - Perdidos 2026",
                "Perdidos 2025"
            ]
        )
        
        df_resumen = resumenes_perd[categoria_sel].copy()
        df_detalle = detalles_perd[categoria_sel].copy()
        
        # Mostrar resumen e histograma
        st.subheader(f"📊 Resumen Temporal: {categoria_sel}")
        
        col_p1, col_p2 = st.columns([3, 4])
        
        with col_p1:
            st.dataframe(
                df_resumen.style.format({
                    'Total Clientes': '{:,}',
                    'Clientes Activos': '{:,}',
                    'Clientes Perdidos': '{:,}',
                    '% Activos': '{:.1%}',
                    '% Perdidos': '{:.1%}'
                }),
                use_container_width=True,
                hide_index=True
            )
            
        with col_p2:
            fig_hist = px.bar(
                df_resumen,
                x='Mes',
                y='Clientes Perdidos',
                title=f'Bajas por Mes (Última Actividad) - {categoria_sel}',
                text_auto=True,
                color_discrete_sequence=['#c65911']
            )
            fig_hist.update_layout(margin=dict(t=30, b=10, l=10, r=10), height=260)
            st.plotly_chart(fig_hist, use_container_width=True)
            
        st.divider()
        
        # Buscador y Detalle
        st.subheader("🔍 Listado y Buscador de Clientes Inactivos")
        
        busqueda_lost = st.text_input(f"Buscar Holder inactivo en '{categoria_sel}':")
        if busqueda_lost:
            df_detalle = df_detalle[df_detalle['Nombre del Holder'].astype(str).str.contains(busqueda_lost, case=False, na=False)]
            
        st.markdown(f"Mostrando **{len(df_detalle)}** de **{len(detalles_perd[categoria_sel])}** holders inactivos.")
        
        st.dataframe(
            df_detalle[['Nombre del Holder', 'Director', 'Última Semana Activo', 'Último Mes Activo', 'Promedio Semanal Histórico']].style.format({
                'Promedio Semanal Histórico': '${:,.2f}'
            }),
            use_container_width=True,
            hide_index=True
        )