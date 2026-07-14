import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from datetime import datetime
import re

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(layout="wide", page_title="Agenda Territorial Ejecutiva", page_icon="🗺️")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    div[data-testid="stForm"] { background-color: white; border-radius: 10px; padding: 30px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.05); }
    h1 { color: #1E3A8A; font-weight: 700; }
    h2 { color: #2C3E50; }
    </style>
""", unsafe_allow_html=True)

# Nombre del archivo Excel que actúa como base de datos
EXCEL_FILE = "agenda_territorial_consolidada.xlsx"

# ==========================================
# 2. FUNCIONES DE LIMPIEZA Y CARGA DE DATOS
# ==========================================

def limpiar_fecha_para_calendario(val):
    """
    Toma fechas del Excel (incluso rangos como '17 y 18/07/2026') 
    y devuelve una fecha limpia en formato YYYY-MM-DD para el calendario.
    """
    val_str = str(val).strip()
    if not val_str or val_str == "nan" or val_str == "":
        return None
    
    # 1. Si ya tiene formato ISO YYYY-MM-DD (con o sin hora)
    match_iso = re.match(r"^(\d{4}-\d{2}-\d{2})", val_str)
    if match_iso:
        return match_iso.group(1)
        
    # Limpiar espacios raros antes/después de las barras diagonales
    val_str = re.sub(r"\s*/\s*", "/", val_str)
        
    # 2. Si es del tipo rango "17 y 18/07/2026" o "16 a 17/07/2026"
    match_rango = re.search(r"(\d+)\s*(?:y|a|-)\s*\d+/(\d+)/(\d{4})", val_str)
    if match_rango:
        dia = match_rango.group(1).zfill(2)
        mes = match_rango.group(2).zfill(2)
        anio = match_rango.group(3)
        return f"{anio}-{mes}-{dia}"
        
    # 3. Si es formato "DD/MM/YYYY" normal
    match_normal = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", val_str)
    if match_normal:
        dia = match_normal.group(1).zfill(2)
        mes = match_normal.group(2).zfill(2)
        anio = match_normal.group(3)
        return f"{anio}-{mes}-{dia}"
        
    return None

def load_data():
    """Carga el Excel, elimina las filas vacías 'fantasmas' y limpia nulos."""
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # --- ELIMINACIÓN DE FILAS FANTASMA ---
        # Si la fila tiene celdas vacías o solo espacios en 'Fecha' y 'Actividad', la descartamos
        df['_temp_fecha'] = df['Fecha'].astype(str).str.strip().replace('', None)
        df['_temp_actividad'] = df['Actividad'].astype(str).str.strip().replace('', None)
        
        df = df[
            (~df['_temp_fecha'].isna() & (df['_temp_fecha'] != 'nan')) |
            (~df['_temp_actividad'].isna() & (df['_temp_actividad'] != 'nan'))
        ]
        df = df.drop(columns=['_temp_fecha', '_temp_actividad'])
        # -------------------------------------

        # Aseguramos que existan todas las columnas requeridas
        columnas_requeridas = ['Fecha', 'Semana', 'Actividad', 'Localidad', 'Organismo/Actor', 'Descripción', 'Estado', 'Público Destinatario', 'Prioridad']
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
                
        # Reemplazar valores nulos reales por textos limpios para evitar que se rompa la vista
        df['Actividad'] = df['Actividad'].fillna("Actividad sin título").astype(str)
        df['Localidad'] = df['Localidad'].fillna("Sin Localidad").astype(str)
        df['Organismo/Actor'] = df['Organismo/Actor'].fillna("No especificado").astype(str)
        df['Descripción'] = df['Descripción'].fillna("").astype(str)
        df['Estado'] = df['Estado'].fillna("Pendiente").astype(str)
        df['Prioridad'] = df['Prioridad'].fillna("INTERMEDIA").astype(str)
        df['Público Destinatario'] = df['Público Destinatario'].fillna("General").astype(str)
        
        return df.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"Error al cargar la base de datos: {e}")
        return pd.DataFrame(columns=['Fecha', 'Semana', 'Actividad', 'Localidad', 'Organismo/Actor', 'Descripción', 'Estado', 'Público Destinatario', 'Prioridad'])

# Inicializar los datos en la sesión de Streamlit
if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data()

# ==========================================
# 3. DISEÑO DE LA INTERFAZ DE USUARIO
# ==========================================
st.title("🗺️ Agenda Territorial Ejecutiva - UPEU")
st.write("Herramienta colaborativa para la carga, modificación y visualización en tiempo real de actividades.")

# Pestañas de navegación
tab1, tab2, tab3, tab4 = st.tabs([
    "🗓️ Vista de Calendario", 
    "✍️ Carga Rápida de Actividad", 
    "✏️ Modificar / Eliminar Actividad",
    "📊 Base de Datos Completa"
])

# ------------------------------------------
# TAB 1: CALENDARIO INTERACTIVO
# ------------------------------------------
with tab1:
    st.header("Planificación Territorial")
    st.write("Haz clic sobre cualquier evento en el calendario para desplegar su ficha de detalles.")
    
    events = []
    colores_prioridad = {
        "ALTA": "#E74C3C",       # Rojo
        "INTERMEDIA": "#F39C12", # Naranja
        "BAJA": "#2ECC71"        # Verde
    }
    
    for idx, row in st.session_state.agenda.iterrows():
        fecha_limpia = limpiar_fecha_para_calendario(row['Fecha'])
        
        if fecha_limpia:
            events.append({
                "title": f"[{row['Localidad']}] {row['Actividad']}",
                "start": fecha_limpia,
                "end": fecha_limpia,
                "color": colores_prioridad.get(str(row['Prioridad']).upper().strip(), "#34495E"),
                "extendedProps": {
                    "fecha_original": str(row['Fecha']),
                    "organismo": str(row['Organismo/Actor']),
                    "descripcion": str(row['Descripción']),
                    "estado": str(row['Estado']),
                    "publico": str(row['Público Destinatario']),
                    "prioridad": str(row['Prioridad'])
                }
            })

    calendar_options = {
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,timeGridWeek,listMonth"
        },
        "initialView": "dayGridMonth",
        "locale": "es",
        "buttonText": {
            "today": "Hoy",
            "month": "Mes",
            "week": "Semana",
            "list": "Lista"
        }
    }
    
    if len(events) > 0:
        state = calendar(events=events, options=calendar_options, key="calendar_agenda")
        
        if state.get("eventClick"):
            clicked = state["eventClick"]["event"]
            props = clicked.get("extendedProps", {})
            
            st.markdown("---")
            st.subheader("🔍 Detalle de la Actividad Seleccionada")
            
            col_det1, col_det2 = st.columns(2)
            with col_det1:
                st.markdown(f"**📌 Actividad:** {clicked['title']}")
                st.markdown(f"**📅 Fecha original:** `{props.get('fecha_original')}`")
                st.markdown(f"**🏢 Organismo/Actor:** {props.get('organismo')}")
                st.markdown(f"**🎯 Público Destinatario:** {props.get('publico')}")
            with col_det2:
                st.markdown(f"**🚨 Prioridad:** `{props.get('prioridad')}`")
                st.markdown(f"**⚙️ Estado:** `{props.get('estado')}`")
                st.markdown(f"**📝 Descripción:** {props.get('descripcion')}")
    else:
        st.warning("No hay eventos programados con fechas válidas para mostrar en el calendario.")

# ------------------------------------------
# TAB 2: FORMULARIO DE CARGA DE DATOS
# ------------------------------------------
with tab2:
    st.header("Registrar Nueva Actividad")
    st.write("Completa el formulario para agregar un nuevo evento en tiempo real:")
    
    with st.form("nuevo_evento_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            f_fecha = st.date_input("Fecha programada", datetime.today())
            f_actividad = st.text_input("Nombre de la Actividad", placeholder="Ej: Entrega de Diplomas")
            f_localidad = st.text_input("Localidad", placeholder="Ej: San Carlos de Bariloche")
            f_organismo = st.text_input("Organismo / Actor principal", placeholder="Ej: Instituto Balseiro")
            
        with col2:
            f_prioridad = st.selectbox("Nivel de Prioridad", ["ALTA", "INTERMEDIA", "BAJA"])
            f_estado = st.selectbox("Estado Actual", ["Pendiente", "En curso", "Finalizado", "Suspendido"])
            f_publico = st.text_input("Público Objetivo", placeholder="Ej: Jóvenes y adultos")
            f_descripcion = st.text_area("Descripción de la acción")
            
        submitted = st.form_submit_button("💾 Guardar en la Agenda")
        
        if submitted:
            if not f_actividad or not f_localidad:
                st.error("Por favor, completa obligatoriamente los campos 'Actividad' y 'Localidad'.")
            else:
                nueva_actividad = {
                    "Fecha": str(f_fecha),
                    "Semana": int(f_fecha.isocalendar()[1]),
                    "Actividad": f_actividad,
                    "Localidad": f_localidad,
                    "Organismo/Actor": f_organismo,
                    "Descripción": f_descripcion,
                    "Estado": f_estado,
                    "Público Destinatario": f_publico,
                    "Prioridad": f_prioridad
                }
                
                st.session_state.agenda = pd.concat([st.session_state.agenda, pd.DataFrame([nueva_actividad])], ignore_index=True)
                st.session_state.agenda.to_excel(EXCEL_FILE, index=False)
                st.success(f"¡Excelente! '{f_actividad}' ha sido guardada.")
                st.rerun()

# ------------------------------------------
# TAB 3: MODIFICAR O ELIMINAR ACTIVIDADES EXISTENTES
# ------------------------------------------
with tab3:
    st.header("Editar / Cancelar Actividades")
    
    df_agenda = st.session_state.agenda.copy()
    
    if len(df_agenda) > 0:
        st.write("Selecciona una actividad de la lista para corregir sus datos o darla de baja:")
        
        # Generar identificadores para el selector
        opciones_actividades = []
        for idx, row in df_agenda.iterrows():
            opciones_actividades.append(f"{idx} | [{row['Localidad']}] {row['Actividad']} ({row['Fecha']})")
            
        actividad_seleccionada = st.selectbox("Seleccionar Actividad a Gestionar", opciones_actividades)
        
        if actividad_seleccionada:
            idx_seleccionado = int(actividad_seleccionada.split(" | ")[0])
            registro_actual = df_agenda.loc[idx_seleccionado]
            
            st.info(f"Modificando el registro de la fila {idx_seleccionado}")
            
            with
