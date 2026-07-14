import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from datetime import datetime

st.set_page_config(layout="wide", page_title="Agenda Territorial Ejecutiva")

# 1. Base de datos simulada (En producción usarías una base de datos o Google Sheets)
# Para este ejemplo, cargamos o iniciamos el archivo
@st.cache_data
def load_data():
    # Aquí puedes leer tu archivo consolidado o inicializar uno vacío
    try:
        df = pd.read_excel("agenda_consolidada.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Fecha', 'Actividad', 'Localidad', 'Organismo_Actor', 'Descripcion', 'Estado', 'Publico_Destinatario', 'Prioridad'])
    return df

if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data()

# Título
st.title("🗺️ Agenda Territorial Ejecutiva")

# Crear pestañas para organizar la herramienta
tab1, tab2, tab3 = st.tabs(["🗓️ Vista de Calendario", "✍️ Cargar Actividad", "📊 Ver Datos y Exportar"])

# --- PESTAÑA 2: CARGA DE DATOS ---
with tab2:
    st.header("Cargar Nueva Actividad Territorial")
    with st.form("form_carga", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            fecha = st.date_input("Fecha del Evento", datetime.today())
            actividad = st.text_input("Nombre de la Actividad", placeholder="Ej: Lanzamiento Programa Desafíos")
            localidad = st.text_input("Localidad", placeholder="Ej: General Roca")
            organismo = st.text_input("Organismo / Actor involucrado", placeholder="Ej: Instituto Balseiro")
        
        with col2:
            estado = st.selectbox("Estado", ["Pendiente", "En curso", "Realizado", "Suspendido"])
            prioridad = st.selectbox("Prioridad", ["ALTA", "INTERMEDIA", "BAJA"])
            publico = st.text_input("Público Destinatario", placeholder="Ej: PyMEs, Estudiantes")
            descripcion = st.text_area("Descripción detallada")
        
        submitted = st.form_submit_submit("Guardar en Agenda")
        if submitted:
            nueva_fila = {
                'Fecha': str(fecha),
                'Actividad': actividad,
                'Localidad': localidad,
                'Organismo_Actor': organismo,
                'Descripcion': descripcion,
                'Estado': estado,
                'Publico_Destinatario': publico,
                'Prioridad': prioridad
            }
            # Guardar en el estado de la app
            st.session_state.agenda = pd.concat([st.session_state.agenda, pd.DataFrame([nueva_fila])], ignore_index=True)
            # Aquí guardarías a tu Excel o Base de Datos física:
            # st.session_state.agenda.to_excel("agenda_consolidada.xlsx", index=False)
            st.success("¡Actividad agendada con éxito!")

# --- PESTAÑA 1: CALENDARIO INTERACTIVO ---
with tab1:
    st.header("Planificación Mensual")
    
    # Preparar los datos del DataFrame para el componente de Calendario
    events = []
    colores_prioridad = {"ALTA": "#FF4B4B", "INTERMEDIA": "#FFAA00", "BAJA": "#00CC96"}
    
    for idx, row in st.session_state.agenda.iterrows():
        if pd.notna(row['Fecha']) and pd.notna(row['Actividad']):
            events.append({
                "title": f"[{row['Localidad']}] {row['Actividad']}",
                "start": str(row['Fecha']),
                "end": str(row['Fecha']),
                "color": colores_prioridad.get(row['Prioridad'], "#3D9970"),
                "extendedProps": {
                    "Organismo": row['Organismo_Actor'],
                    "Estado": row['Estado'],
                    "Prioridad": row['Prioridad']
                }
            })
            
    calendar_options = {
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,timeGridWeek"
        },
        "initialView": "dayGridMonth",
        "selectable": True,
    }
    
    # Renderizar el calendario interactivo
    state = calendar(events=events, options=calendar_options, key="agenda_calendar")
    
    # Mostrar detalles si hacen clic en un evento
    if state.get("eventClick"):
        event_clicked = state["eventClick"]["event"]
        st.info(f"**Detalles de la Actividad Seleccionada:**\n"
                f"- **Título:** {event_clicked['title']}\n"
                f"- **Fecha:** {event_clicked['start']}\n"
                f"- **Estado:** {event_clicked['extendedProps'].get('Estado')}\n"
                f"- **Prioridad:** {event_clicked['extendedProps'].get('Prioridad')}")

# --- PESTAÑA 3: TABLA DE DATOS Y FILTROS ---
with tab3:
    st.header("Base de Datos Completa")
    # Filtros rápidos
    localidades_disponibles = ["Todas"] + list(st.session_state.agenda['Localidad'].dropna().unique())
    filtro_loc = st.selectbox("Filtrar por Localidad", localidades_disponibles)
    
    df_filtrado = st.session_state.agenda.copy()
    if filtro_loc != "Todas":
        df_filtrado = df_filtrado[df_filtrado['Localidad'] == filtro_loc]
        
    st.dataframe(df_filtrado, use_container_width=True)
    
    # Botón para descargar de nuevo a un Excel limpio
    excel_data = df_filtrado.to_excel(index=False) # Requiere lógica de conversión a bytes
    st.download_button("Descargar Planificación Filtrada (Excel)", data="", file_name="agenda_territorial_export.xlsx")