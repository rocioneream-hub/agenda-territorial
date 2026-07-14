import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from datetime import datetime
import re
import os

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS (IDENTIDAD VISUAL OFICIAL)
# ==========================================
st.set_page_config(
    layout="wide", 
    page_title="Portal de Gestión - Gobierno de Río Negro", 
    page_icon="📈"
)

# Inyección de estilos CSS basados estrictamente en el Manual de Marca de Río Negro
st.markdown("""
    <style>
    /* Importación de la tipografía oficial Figtree */
    @import url('https://fonts.googleapis.com/css2?family=Figtree:ital,wght@0,300..900;1,300..900&display=swap');
    
    /* Configuración de fuentes globales */
    html, body, [class*="css"], .stMarkdown, p, div {
        font-family: 'Figtree', sans-serif !important;
    }
    
    /* Fondo general utilizando estrictamente el Gris RN oficial (#E8E8E8) */
    .stApp, .main, [data-testid="stAppViewContainer"] { 
        background-color: #E8E8E8 !important; 
    }
    
    /* Diseño de tarjetas y formularios (Fondo blanco rígido y bordes limpios) */
    div[data-testid="stForm"] { 
        background-color: #FFFFFF !important; 
        border-radius: 8px !important; 
        padding: 30px !important; 
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.05) !important;
        border: 1px solid #D1D5DB !important;
    }
    
    /* Títulos Principales en Negro Puro */
    h1 { 
        color: #000000 !important; 
        font-weight: 800 !important;
        font-size: 2.2rem !important;
        margin-top: 10px !important;
        margin-bottom: 0px !important;
    }
    
    /* Subtítulos en el Azul RN oficial (#007BE0) */
    h2, h3 { 
        color: #007BE0 !important; 
        font-weight: 700 !important;
    }
    
    /* Estilo del botón primario con el Verde RN oficial (#6AC64F) */
    button[kind="primary"] {
        background-color: #6AC64F !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        border: none !important;
        transition: background-color 0.2s ease;
    }
    
    button[kind="primary"]:hover {
        background-color: #59b040 !important; /* Verde RN ligeramente más oscuro para el hover */
    }
    
    /* Hashtag de gestión oficial #gobiernodelosrionegrinos */
    .hashtag-gestion {
        color: #6AC64F !important;
        font-weight: 800;
        font-size: 1.1rem;
    }

    /* ==========================================
       PERSONALIZACIÓN CSS PARA EL CALENDARIO (FULLCALENDAR)
       ========================================== */
    
    /* Botones estándar de navegación y de vista del calendario (Estilo Azul RN) */
    .fc .fc-button-primary {
        background-color: #007BE0 !important;
        border-color: #007BE0 !important;
        color: #FFFFFF !important;
        font-family: 'Figtree', sans-serif !important;
        font-weight: 600 !important;
        text-transform: capitalize !important;
        box-shadow: none !important;
        transition: background-color 0.2s ease, border-color 0.2s ease !important;
    }

    /* Efecto Hover en los botones del calendario (Verde RN) */
    .fc .fc-button-primary:hover {
        background-color: #6AC64F !important;
        border-color: #6AC64F !important;
    }

    /* Estado de botón activo o seleccionado (Ej: vista "Mes" seleccionada) */
    .fc .fc-button-primary:not(:disabled).fc-button-active, 
    .fc .fc-button-primary:not(:disabled):active {
        background-color: #00569e !important; /* Azul RN oscuro para el estado activo */
        border-color: #00569e !important;
    }

    /* Botón deshabilitado del calendario */
    .fc .fc-button-primary:disabled {
        background-color: #b3d7f5 !important;
        border-color: #b3d7f5 !important;
        color: #FFFFFF !important;
    }

    /* Personalización de los encabezados de los días en el calendario */
    .fc .fc-col-header-cell-cushion {
        color: #000000 !important;
        font-weight: 700 !important;
        text-decoration: none !important;
    }
    
    </style>
""", unsafe_allow_html=True)

# Nombre del archivo Excel que actúa como base de datos y logo personalizado actualizado
EXCEL_FILE = "agenda_territorial_consolidada.xlsx"
LOGO_FILE = "isologo_RN.png"

# ==========================================
# 2. CONTROL DE ACCESO (MODO LECTOR / EDITOR)
# ==========================================
st.sidebar.header("🔑 Control de Acceso")
password = st.sidebar.text_input("Contraseña de Editor", type="password")

# Contraseña de seguridad para la edición
CONTRASEÑA_CORRECTA = "UPEU2026" 

es_editor = (password == CONTRASEÑA_CORRECTA)

if es_editor:
    st.sidebar.success("🔑 Modo Editor Activado")
else:
    st.sidebar.info("👁️ Modo Visualización (Solo Lectura)")

# ==========================================
# 3. FUNCIONES DE LIMPIEZA Y CARGA DE DATOS
# ==========================================

def limpiar_fecha_para_calendario(val):
    val_str = str(val).strip()
    if not val_str or val_str == "nan" or val_str == "":
        return None
    
    # 1. Formato estándar YYYY-MM-DD
    match_iso = re.match(r"^(\d{4}-\d{2}-\d{2})", val_str)
    if match_iso:
        return match_iso.group(1)
        
    val_str = re.sub(r"\s*/\s*", "/", val_str)
        
    # 2. Formato rango "17 y 18/07/2026"
    match_rango = re.search(r"(\d+)\s*(?:y|a|-)\s*\d+/(\d+)/(\d{4})", val_str)
    if match_rango:
        dia = match_rango.group(1).zfill(2)
        mes = match_rango.group(2).zfill(2)
        anio = match_rango.group(3)
        return f"{anio}-{mes}-{dia}"
        
    # 3. Formato DD/MM/YYYY normal
    match_normal = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", val_str)
    if match_normal:
        dia = match_normal.group(1).zfill(2)
        mes = match_normal.group(2).zfill(2)
        anio = match_normal.group(3)
        return f"{anio}-{mes}-{dia}"
        
    return None

def load_data():
    """Carga el Excel, normaliza nombres de columnas, elimina filas fantasma y limpia nulos."""
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # Normalizar los nombres de las columnas para quitar espacios accidentales
        df.columns = df.columns.str.strip()
        
        # --- ELIMINACIÓN DE FILAS FANTASMA ---
        df['_temp_fecha'] = df['Fecha'].astype(str).str.strip().replace('', None)
        df['_temp_actividad'] = df['Actividad'].astype(str).str.strip().replace('', None)
        
        df = df[
            (~df['_temp_fecha'].isna() & (df['_temp_fecha'] != 'nan') & (df['_temp_fecha'] != '')) |
            (~df['_temp_actividad'].isna() & (df['_temp_actividad'] != 'nan') & (df['_temp_actividad'] != ''))
        ]
        df = df.drop(columns=['_temp_fecha', '_temp_actividad'])

        # Asegurar que existan todas las columnas requeridas
        columnas_requeridas = [
            'Fecha', 'Semana', 'Actividad', 'Localidad', 'Organismo/Actor', 
            'Descripción', 'Estado', 'Público Destinatario', 'Prioridad', 'Invitación a participar'
        ]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
                
        df['Actividad'] = df['Actividad'].fillna("Actividad sin título").astype(str)
        df['Localidad'] = df['Localidad'].fillna("Sin Localidad").astype(str).str.strip()
        df['Organismo/Actor'] = df['Organismo/Actor'].fillna("No especificado").astype(str)
        df['Descripción'] = df['Descripción'].fillna("").astype(str)
        df['Estado'] = df['Estado'].fillna("Pendiente").astype(str)
        df['Prioridad'] = df['Prioridad'].fillna("INTERMEDIA").astype(str)
        df['Público Destinatario'] = df['Público Destinatario'].fillna("General").astype(str)
        df['Invitación a participar'] = df['Invitación a participar'].fillna("").astype(str)
        
        return df.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"Error al cargar la base de datos: {e}")
        return pd.DataFrame(columns=[
            'Fecha', 'Semana', 'Actividad', 'Localidad', 'Organismo/Actor', 
            'Descripción', 'Estado', 'Público Destinatario', 'Prioridad', 'Invitación a participar'
        ])

# Inicializar o forzar la carga de la agenda
if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data()

# ==========================================
# 4. DISEÑO DE LA INTERFAZ DE USUARIO (LOGO ARRIBA DEL TODO)
# ==========================================

# El logo se muestra con un tamaño mediano y elegante de 180px para evitar pixelaciones
if os.path.exists(LOGO_FILE):
    st.image(LOGO_FILE, width=180)
else:
    st.info("Logotipo Río Negro")

st.markdown("---")  # Línea sutil divisoria debajo del logo oficial

# Cabecera de títulos y aplicación
col_title_left, col_title_right = st.columns([4, 1.5])

with col_title_left:
    st.title("Agenda de Planificación Territorial")
    st.markdown("**Unidad Provincial de Enlace con Universidades (UPEU)** | Gobierno de Río Negro")
    st.markdown("<span class='hashtag-gestion'>#gobiernodelosrionegrinos</span>", unsafe_allow_html=True) 

with col_title_right:
    st.write("")
    st.write("")
    if st.button("🔄 Sincronizar Excel", use_container_width=True):
        st.session_state.agenda = load_data()
        st.success("¡Base de datos sincronizada!")
        st.rerun()

# Definición dinámica de pestañas dependiendo de si el usuario ingresó la contraseña
if es_editor:
    tab1, tab2, tab3, tab4 = st.tabs([
        "🗓️ Vista de Calendario", 
        "✍️ Carga Rápida de Actividad", 
        "✏️ Modificar / Eliminar Actividad",
        "📊 Base de Datos Completa"
    ])
else:
    tab1, tab4 = st.tabs([
        "🗓️ Vista de Calendario", 
        "📊 Base de Datos Completa"
    ])
    tab2, tab3 = None, None

# ------------------------------------------
# TAB 1: CALENDARIO INTERACTIVO (Disponible para todos)
# ------------------------------------------
with tab1:
    st.header("Planificación Territorial")
    st.write("Haz clic sobre cualquier evento en el calendario para desplegar su ficha de detalles.")
    
    events = []
    # Paleta de colores oficial por Prioridad
    colores_prioridad = {
        "ALTA": "#E74C3C",       # Rojo
        "INTERMEDIA": "#F39C12", # Naranja
        "BAJA": "#6AC64F"        # Verde RN Oficial
    }
    
    # Color especial Azul RN oficial para destacar eventos con Invitación formal a participar
    COLOR_CON_INVITACION = "#007BE0" 
    
    for idx, row in st.session_state.agenda.iterrows():
        fecha_limpia = limpiar_fecha_para_calendario(row['Fecha'])
        
        if fecha_limpia:
            # Detectamos si tiene cargado algo en el campo de invitación
            invitacion_val = str(row.get('Invitación a participar', '')).strip()
            tiene_invitacion = invitacion_val != "" and invitacion_val.lower() != "nan"
            
            # Si tiene invitación, asignamos el azul RN y el emoji de sobre.
            if tiene_invitacion:
                color_evento = COLOR_CON_INVITACION
                titulo_mostrar = f"✉️ [{row['Localidad']}] {row['Actividad']}"
            else:
                color_evento = colores_prioridad.get(str(row['Prioridad']).upper().strip(), "#000000")
                titulo_mostrar = f"[{row['Localidad']}] {row['Actividad']}"
                
            events.append({
                "title": titulo_mostrar,
                "start": fecha_limpia,
                "end": fecha_limpia,
                "color": color_evento,
                "extendedProps": {
                    "fecha_original": str(row['Fecha']),
                    "organismo": str(row['Organismo/Actor']),
                    "descripcion": str(row['Descripción']),
                    "estado": str(row['Estado']),
                    "publico": str(row['Público Destinatario']),
                    "prioridad": str(row['Prioridad']),
                    "invitacion": invitacion_val if tiene_invitacion else "Sin invitaciones especiales"
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
                st.markdown(f"**✉️ Invitación a participar:** `{props.get('invitacion')}`")
                st.markdown(f"**📝 Descripción:** {props.get('descripcion')}")
    else:
        st.warning("No hay eventos programados con fechas válidas para mostrar en el calendario.")

# ------------------------------------------
# TAB 2: FORMULARIO DE CARGA DE DATOS (Solo visible para Editores)
# ------------------------------------------
if es_editor and tab2 is not None:
    with tab2:
        st.header("Registrar Nueva Actividad")
        st.write("Completa el formulario para agregar un nuevo evento en tiempo real:")
        
        with st.form("nuevo_evento_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                f_fecha = st.date_input("Fecha programada", datetime.today())
                f_actividad = st.text_input("Nombre de la Actividad", placeholder="Ej: Lanzamiento Programa Desafíos")
                f_localidad = st.text_input("Localidad", placeholder="Ej: General Roca")
                f_organismo = st.text_input("Organismo / Actor principal", placeholder="Ej: Secretaría de Estado de Energía y Ambiente")
                
            with col2:
                f_prioridad = st.selectbox("Nivel de Prioridad", ["ALTA", "INTERMEDIA", "BAJA"])
                f_estado = st.selectbox("Estado Actual", ["Pendiente", "En curso", "Finalizado", "Suspendido"])
                f_publico = st.text_input("Público Objetivo", placeholder="Ej: Jóvenes y adultos")
                f_invitacion = st.text_input("Invitación a participar", placeholder="Ej: Gobernador + Secretaria de EEYA")
                f_descripcion = st.text_area("Descripción de la acción / Notas")
                
            submitted = st.form_submit_button("💾 Guardar en la Agenda")
            
            if submitted:
                if not f_actividad or not f_localidad:
                    st.error("Por favor, completa obligatoriamente los campos 'Actividad' y 'Localidad'.")
                else:
                    nueva_actividad = {
                        "Fecha": str(f_fecha),
                        "Semana": int(f_fecha.isocalendar()[1]),
                        "Actividad": f_actividad,
                        "Localidad": f_localidad.strip(),
                        "Organismo/Actor": f_organismo,
                        "Descripción": f_descripcion,
                        "Estado": f_estado,
                        "Público Destinatario": f_publico,
                        "Prioridad": f_prioridad,
                        "Invitación a participar": f_invitacion
                    }
                    
                    st.session_state.agenda = pd.concat([st.session_state.agenda, pd.DataFrame([nueva_actividad])], ignore_index=True)
                    st.session_state.agenda.to_excel(EXCEL_FILE, index=False)
                    st.success(f"¡Excelente! '{f_actividad}' ha sido guardada exitosamente.")
                    st.rerun()

# ------------------------------------------
# TAB 3: MODIFICAR O ELIMINAR ACTIVIDADES (Solo visible para Editores)
# ------------------------------------------
if es_editor and tab3 is not None:
    with tab3:
        st.header("Editar / Cancelar Actividades")
        
        df_agenda = st.session_state.agenda.copy()
        
        if len(df_agenda) > 0:
            st.write("Selecciona una actividad de la lista para corregir sus datos o darla de baja:")
            
            opciones_actividades = []
            for idx, row in df_agenda.iterrows():
                opciones_actividades.append(f"{idx} | [{row['Localidad']}] {row['Actividad']} ({row['Fecha']})")
                
            actividad_seleccionada = st.selectbox("Seleccionar Actividad a Gestionar", opciones_actividades)
            
            if actividad_seleccionada:
                idx_seleccionado = int(actividad_seleccionada.split(" | ")[0])
                registro_actual = df_agenda.loc[idx_seleccionado]
                
                st.info(f"Modificando el registro de la fila {idx_seleccionado}")
                
                with st.form("form_edicion"):
                    col1_ed, col2_ed = st.columns(2)
                    
                    with col1_ed:
                        try:
                            fecha_previa = datetime.strptime(limpiar_fecha_para_calendario(registro_actual['Fecha']), "%Y-%m-%d").date()
                        except:
                            fecha_previa = datetime.today().date()
                            
                        ed_fecha = st.date_input("Fecha", value=fecha_previa)
                        ed_actividad = st.text_input("Nombre de la Actividad", value=str(registro_actual['Actividad']))
                        ed_localidad = st.text_input("Localidad", value=str(registro_actual['Localidad']))
                        ed_organismo = st.text_input("Organismo / Actor", value=str(registro_actual['Organismo/Actor']))
                        
                    with col2_ed:
                        opciones_prioridad = ["ALTA", "INTERMEDIA", "BAJA"]
                        def_prio = opciones_prioridad.index(str(registro_actual['Prioridad']).upper().strip()) if str(registro_actual['Prioridad']).upper().strip() in opciones_prioridad else 1
                        
                        opciones_estado = ["Pendiente", "En curso", "Finalizado", "Suspendido"]
                        def_est = opciones_estado.index(str(registro_actual['Estado']).capitalize().strip()) if str(registro_actual['Estado']).capitalize().strip() in opciones_estado else 0
                        
                        ed_prioridad = st.selectbox("Prioridad", opciones_prioridad, index=def_prio)
                        ed_estado = st.selectbox("Estado", opciones_estado, index=def_est)
                        ed_publico = st.text_input("Público Destinatario", value=str(registro_actual['Público Destinatario']))
                        ed_invitacion = st.text_input("Invitación a participar", value=str(registro_actual.get('Invitación a participar', '')))
                        ed_descripcion = st.text_area("Descripción", value=str(registro_actual['Descripción']))
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        boton_actualizar = st.form_submit_button("🔄 Actualizar Cambios")
                    with col_btn2:
                        boton_eliminar = st.form_submit_button("❌ Eliminar Actividad permanentemente")
                        
                    if boton_actualizar:
                        st.session_state.agenda.at[idx_seleccionado, 'Fecha'] = str(ed_fecha)
                        st.session_state.agenda.at[idx_seleccionado, 'Semana'] = int(ed_fecha.isocalendar()[1])
                        st.session_state.agenda.at[idx_seleccionado, 'Actividad'] = ed_actividad
                        st.session_state.agenda.at[idx_seleccionado, 'Localidad'] = ed_localidad.strip()
                        st.session_state.agenda.at[idx_seleccionado, 'Organismo/Actor'] = ed_organismo
                        st.session_state.agenda.at[idx_seleccionado, 'Descripción'] = ed_descripcion
                        st.session_state.agenda.at[idx_seleccionado, 'Estado'] = ed_estado
                        st.session_state.agenda.at[idx_seleccionado, 'Público Destinatario'] = ed_publico
                        st.session_state.agenda.at[idx_seleccionado, 'Prioridad'] = ed_prioridad
                        st.session_state.agenda.at[idx_seleccionado, 'Invitación a participar'] = ed_invitacion
                        
                        st.session_state.agenda.to_excel(EXCEL_FILE, index=False)
                        st.success("¡Actividad actualizada con éxito!")
                        st.rerun()
                        
                    if boton_eliminar:
                        st.session_state.agenda = st.session_state.agenda.drop(idx_seleccionado).reset_index(drop=True)
                        st.session_state.agenda.to_excel(EXCEL_FILE, index=False)
                        st.warning("La actividad ha sido eliminada del sistema.")
                        st.rerun()
        else:
            st.warning("No hay actividades registradas en la base de datos para editar.")

# ------------------------------------------
# TAB 4: TABLA DE DATOS, BUSCADOR Y EXPORTACIÓN (Disponible para todos)
# ------------------------------------------
with tab4:
    st.header("Buscador y Reportes")
    
    df_filtrado = st.session_state.agenda.copy()
    
    if len(df_filtrado) > 0:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            search_query = st.text_input("Buscar por palabra clave...", key="search_tab4").lower()
        with col_f2:
            localidades_unicas = sorted(list(df_filtrado['Localidad'].astype(str).str.strip().unique()))
            list_localidades = ["Todas"] + [loc for loc in localidades_unicas if loc != 'nan' and loc != '']
            filtro_localidad = st.selectbox("Filtrar por Localidad", list_localidades, key="filter_loc_tab4")
            
        # Filtrar localidad
        if filtro_localidad != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Localidad'].str.strip() == filtro_localidad]
            
        # Filtrar palabra clave
        if search_query:
            df_filtrado = df_filtrado[
                df_filtrado['Actividad'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Descripción'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Organismo/Actor'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Invitación a participar'].astype(str).str.lower().str.contains(search_query)
            ]
            
        # Renderizado de la tabla
        st.dataframe(df_filtrado, use_container_width=True)
        
        # Exportador a Excel seguro
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Agenda')
        excel_bytes = output.getvalue()
        
        st.download_button(
            label="📥 Descargar datos filtrados a Excel",
            data=excel_bytes,
            file_name="agenda_territorial_filtrada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_tab4"
        )
    else:
        st.warning("La base de datos está vacía o cargando. Prueba pulsar el botón '🔄 Sincronizar Excel' arriba a la derecha.")
