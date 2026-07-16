import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from datetime import datetime, time
import re
import os
import io

# Importación segura de la librería python-docx
try:
    import docx
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    LIBRERIA_DOCX_LISTA = True
except ImportError:
    LIBRERIA_DOCX_LISTA = False

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS (IDENTIDAD VISUAL OFICIAL)
# ==========================================
st.set_page_config(
    layout="wide", 
    page_title="Calendario territorial UPEU", 
    page_icon="🗓️"
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
    .stApp, .main, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stMainBlockContainer"] { 
        background-color: #E8E8E8 !important; 
        background: #E8E8E8 !important;
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

    /* Forzar nitidez absoluta en el renderizado de imágenes vectoriales */
    img {
        image-rendering: -webkit-optimize-contrast !important;
        image-rendering: crisp-edges !important;
    }

    /* ==========================================
       ANULAR EL ROJO NATIVO DE STREAMLIT (TABS Y SELECCIONES)
       ========================================== */
    
    /* Cambiar la línea roja inferior de las Pestañas Activas a Azul RN */
    div[data-baseweb="tab-list"] button[aria-selected="true"] {
        color: #007BE0 !important;
        border-bottom-color: #007BE0 !important;
    }
    
    /* Cambiar el texto de las pestañas no seleccionadas a negro */
    div[data-baseweb="tab-list"] button[aria-selected="false"] {
        color: #333333 !important;
    }

    /* Cambiar el borde de enfoque en inputs y dropdowns a Azul RN */
    .stTextInput input:focus, 
    .stSelectbox div[role="button"]:focus, 
    .stTextArea textarea:focus,
    div[data-baseweb="select"] > div:focus-within {
        border-color: #007BE0 !important;
        box-shadow: 0 0 0 1px #007BE0 !important;
    }

    /* Cambiar el color de los Spinners de carga a Verde RN */
    div[data-testid="stSpinner"] > div {
        border-top-color: #6AC64F !important;
    }

    /* ==========================================
       PERSONALIZACIÓN CSS PARA EL CALENDARIO (FULLCALENDAR)
       ========================================== */
    
    /* Forzar color Azul RN en botones nativos del calendario */
    .fc .fc-button,
    .fc .fc-button-primary,
    .fc-button,
    .fc-button-primary,
    button.fc-button,
    button.fc-today-button,
    button.fc-prev-button,
    button.fc-next-button {
        background-color: #007BE0 !important;
        background: #007BE0 !important;
        border-color: #007BE0 !important;
        color: #FFFFFF !important;
        opacity: 1 !important;
        box-shadow: none !important;
        font-family: 'Figtree', sans-serif !important;
        font-weight: 600 !important;
        text-transform: capitalize !important;
        transition: background-color 0.2s ease, border-color 0.2s ease !important;
    }

    /* Hover en los botones (Verde RN) */
    .fc .fc-button:hover,
    .fc .fc-button-primary:hover,
    .fc-button:hover,
    .fc-button-primary:hover,
    button.fc-button:hover,
    button.fc-today-button:hover,
    button.fc-prev-button:hover,
    button.fc-next-button:hover {
        background-color: #6AC64F !important;
        background: #6AC64F !important;
        border-color: #6AC64F !important;
        color: #FFFFFF !important;
    }

    /* Botón activo seleccionado (Azul RN Oscuro) */
    .fc .fc-button-primary:not(:disabled).fc-button-active, 
    .fc .fc-button-primary:not(:disabled):active,
    .fc-button-active,
    button.fc-button-active,
    .fc .fc-button-primary:active {
        background-color: #00569E !important;
        background: #00569E !important;
        border-color: #00569E !important;
        color: #FFFFFF !important;
    }

    /* Cabeceras de los días del calendario */
    .fc .fc-col-header-cell-cushion {
        color: #000000 !important;
        font-weight: 700 !important;
        text-decoration: none !important;
    }

    .fc-event, .fc-event-dot {
        border-color: transparent !important;
    }

    /* Día de hoy (celeste muy sutil y número azul) */
    .fc .fc-daygrid-day.fc-day-today {
        background-color: rgba(0, 123, 224, 0.08) !important;
    }
    .fc .fc-day-today .fc-daygrid-day-number {
        color: #007BE0 !important;
        font-weight: 800 !important;
    }
    .fc .fc-day-today .fc-daygrid-day-top {
        border-top: 3px solid #007BE0 !important;
    }

    /* Indicador de hora actual en Verde RN */
    .fc .fc-timegrid-now-indicator-line {
        border-color: #6AC64F !important;
        border-width: 2px !important;
    }
    .fc .fc-timegrid-now-indicator-arrow {
        border-top-color: #6AC64F !important;
        border-bottom-color: #6AC64F !important;
    }
    
    </style>
""", unsafe_allow_html=True)

# Nombre del archivo Excel de base de datos e ISOLOGO vectorial
EXCEL_FILE = "agenda_territorial_consolidada.xlsx"
LOGO_FILE = "isologo_RN.svg"

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
            'Fecha', 'Hora', 'Semana', 'Actividad', 'Ciudad', 'Lugar',
            'Explicación breve de la actividad', 'Cantidad de personas estimadas',
            'Organismo/Actor', 'Estado', 'Público Destinatario', 'Prioridad', 'Invitación a participar'
        ]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
                
        df['Actividad'] = df['Actividad'].fillna("Actividad sin título").astype(str)
        df['Hora'] = df['Hora'].fillna("").astype(str).str.strip()
        df['Ciudad'] = df['Ciudad'].fillna("Sin Ciudad").astype(str).str.strip()
        df['Lugar'] = df['Lugar'].fillna("Sin especificar").astype(str).str.strip()
        df['Explicación breve de la actividad'] = df['Explicación breve de la actividad'].fillna("").astype(str)
        
        # Validación segura de asistencia estimada
        df['Cantidad de personas estimadas'] = pd.to_numeric(df['Cantidad de personas estimadas'], errors='coerce').fillna(0).astype(int)
        
        df['Organismo/Actor'] = df['Organismo/Actor'].fillna("No especificado").astype(str)
        df['Estado'] = df['Estado'].fillna("Pendiente").astype(str)
        df['Prioridad'] = df['Prioridad'].fillna("INTERMEDIA").astype(str)
        df['Público Destinatario'] = df['Público Destinatario'].fillna("General").astype(str)
        df['Invitación a participar'] = df['Invitación a participar'].fillna("").astype(str)
        
        return df.reset_index(drop=True)
        
    except Exception as e:
        st.error(f"Error al cargar la base de datos: {e}")
        return pd.DataFrame(columns=[
            'Fecha', 'Hora', 'Semana', 'Actividad', 'Ciudad', 'Lugar',
            'Explicación breve de la actividad', 'Cantidad de personas estimadas',
            'Organismo/Actor', 'Estado', 'Público Destinatario', 'Prioridad', 'Invitación a participar'
        ])

# Inicializar o forzar la carga de la agenda
if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data()

# ==========================================
# FUNCIONES AUXILIARES DE EXPORTACIÓN A WORD (CON FILTROS DE SEGURIDAD EXTREMOS)
# ==========================================

def set_cell_background(cell, color_hex):
    """Establece de manera segura el color de fondo de una celda en una tabla de Word."""
    try:
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))
    except:
        pass

def crear_reporte_word_areas(df):
    """Genera un reporte DOCX formal y tolerante a cualquier tipo de error o columna faltante."""
    doc = docx.Document()
    
    # Configuración de márgenes estándar
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # 1. MEMBRETADO INSTITUCIONAL
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_gob = p_header.add_run("GOBIERNO DE LA PROVINCIA DE RÍO NEGRO\n")
    run_gob.font.bold = True
    run_gob.font.size = Pt(10)
    run_gob.font.color.rgb = docx.shared.RGBColor(106, 198, 79) # Verde RN (#6AC64F)
    
    run_sub = p_header.add_run("Ministerio de Educación y Derechos Humanos\nUnidad Provincial de Enlace con Universidades (UPEU)\n")
    run_sub.font.size = Pt(9.5)
    run_sub.font.color.rgb = docx.shared.RGBColor(100, 100, 100)
    
    # Línea divisoria decorativa
    p_line = doc.add_paragraph()
    p_line_border = OxmlElement('w:pBdr')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), '8')
    bottom_border.set(qn('w:space'), '1')
    bottom_border.set(qn('w:color'), '007BE0') # Azul RN
    p_line_border.append(bottom_border)
    p_line._p.get_or_add_pPr().append(p_line_border)
    
    # 2. TÍTULOS PRINCIPALES
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_title = p_title.add_run("REPORTE EJECUTIVO DE ACTIVIDADES TERRITORIALES")
    run_title.font.bold = True
    run_title.font.size = Pt(16)
    run_title.font.color.rgb = docx.shared.RGBColor(0, 123, 224) # Azul RN
    p_title.paragraph_format.space_after = Pt(2)
    
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_date = p_date.add_run(f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')} | Coordinación Interinstitucional")
    run_date.font.italic = True
    run_date.font.size = Pt(9.5)
    p_date.paragraph_format.space_after = Pt(24)
    
    # 3. CUADRO RESUMEN DE ASISTENCIA GLOBAL (Protección si la columna de asistencia es nula)
    total_acciones = len(df)
    try:
        total_personas = df['Cantidad de personas estimadas'].fillna(0).astype(int).sum()
    except:
        total_personas = 0
    
    table_resumen = doc.add_table(rows=1, cols=2)
    hdr_res = table_resumen.rows[0].cells
    hdr_res[0].text = "Total de Acciones Planificadas"
    hdr_res[1].text = "Asistencia Estimada Global"
    
    for cell in hdr_res:
        set_cell_background(cell, "F5F5F5")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(10)
        p.runs[0].font.color.rgb = docx.shared.RGBColor(51, 51, 51)
        
    row_res = table_resumen.add_row().cells
    row_res[0].text = str(total_acciones)
    row_res[1].text = f"{total_personas:,} personas"
    
    for cell in row_res:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(12)
        p.runs[0].font.color.rgb = docx.shared.RGBColor(0, 123, 224)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # 4. DESGLOSE DE LAS ACTIVIDADES
    doc.add_heading("Fichas de Planificación de Actividades", level=2)
    
    if total_acciones > 0:
        # Aseguramos un orden temporal prolijo
        df_ordenado = df.copy()
        if 'Fecha' in df_ordenado.columns:
            df_ordenado = df_ordenado.sort_values(by='Fecha').reset_index(drop=True)
        
        for idx, row in df_ordenado.iterrows():
            p_act = doc.add_paragraph()
            p_act.paragraph_format.space_before = Pt(14)
            p_act.paragraph_format.space_after = Pt(4)
            
            act_titulo = str(row.get('Actividad', 'Sin Nombre'))
            run_header = p_act.add_run(f"📌 Actividad {idx+1}: {act_titulo}")
            run_header.font.bold = True
            run_header.font.size = Pt(11)
            run_header.font.color.rgb = docx.shared.RGBColor(0, 123, 224)
            
            # Tabla de Ficha Técnica Individual (6 Campos Requeridos)
            ficha_table = doc.add_table(rows=6, cols=2)
            ficha_table.style = 'Light Shading Accent 1'
            ficha_table.columns[0].width = Inches(2.2)
            ficha_table.columns[1].width = Inches(4.3)
            
            # Campo 1: Actividad
            ficha_table.rows[0].cells[0].text = "Actividad:"
            ficha_table.rows[0].cells[1].text = act_titulo
            
            # Campo 2: Ciudad
            ficha_table.rows[1].cells[0].text = "Ciudad / Localidad:"
            ficha_table.rows[1].cells[1].text = str(row.get('Ciudad', 'Sin especificar'))
            
            # Campo 3: Fecha / Hora
            fecha_val = str(row.get('Fecha', 'Sin fecha'))
            hora_val = str(row.get('Hora', '')).strip()
            hora_text = f" - {hora_val} hs" if hora_val else ""
            ficha_table.rows[2].cells[0].text = "Fecha y Horario:"
            ficha_table.rows[2].cells[1].text = f"{fecha_val}{hora_text}"
            
            # Campo 4: Lugar
            ficha_table.rows[3].cells[0].text = "Lugar / Espacio Físico:"
            ficha_table.rows[3].cells[1].text = str(row.get('Lugar', 'Sin especificar'))
            
            # Campo 5: Explicación breve
            exp_val = str(row.get('Explicación breve de la actividad', row.get('Descripción', 'Sin notas adicionales')))
            ficha_table.rows[4].cells[0].text = "Explicación breve:"
            ficha_table.rows[4].cells[1].text = exp_val
            
            # Campo 6: Cantidad de personas estimadas
            try:
                asistencia_num = int(row.get('Cantidad de personas estimadas', 0))
            except:
                asistencia_num = 0
            ficha_table.rows[5].cells[0].text = "Asistencia estimada:"
            ficha_table.rows[5].cells[1].text = f"{asistencia_num:,} asistentes"
            
            # Formatear estilos de celdas de la tabla
            for r_idx, r in enumerate(ficha_table.rows):
                r.cells[0].paragraphs[0].runs[0].font.bold = True
                r.cells[0].paragraphs[0].runs[0].font.size = Pt(9.5)
                r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)
                if r_idx == 5:
                    r.cells[1].paragraphs[0].runs[0].font.bold = True
                    r.cells[1].paragraphs[0].runs[0].font.color.rgb = docx.shared.RGBColor(106, 198, 79) # Verde RN
            
            doc.add_paragraph().paragraph_format.space_after = Pt(6)
    else:
        doc.add_paragraph("No hay actividades cargadas para el reporte.")
        
    # 5. PIE DE PÁGINA CON EL HASHTAG DE GESTIÓN
    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.paragraph_format.space_before = Pt(30)
    run_hashtag = p_footer.add_run("#gobiernodelosrionegrinos")
    run_hashtag.font.bold = True
    run_hashtag.font.size = Pt(11)
    run_hashtag.font.color.rgb = docx.shared.RGBColor(106, 198, 79) # Verde RN
    
    # Guardar documento en buffer de memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 4. DISEÑO DE LA INTERFAZ DE USUARIO (LOGO ARRIBA DEL TODO)
# ==========================================

# El logo se renderiza en SVG con un ancho elegante de 180px para máxima definición
if os.path.exists(LOGO_FILE):
    st.image(LOGO_FILE, width=180)
else:
    st.info("Logotipo Río Negro (SVG)")

st.markdown("---")  # Línea divisoria bajo el logo oficial

# Cabecera de títulos de la app
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

# Pestañas condicionales según el rol del usuario (Editor / Lector)
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
    
    # PALETA DE COLORES ADAPTADA:
    colores_prioridad = {
        "ALTA": "#007BE0",       # Azul RN Oficial
        "INTERMEDIA": "#333333", # Gris Carbón
        "BAJA": "#6AC64F"        # Verde RN Oficial
    }
    
    # Color Naranja Vibrante para destacar invitaciones especiales de manera inequívoca
    COLOR_CON_INVITACION = "#FF7A00" 
    
    for idx, row in st.session_state.agenda.iterrows():
        fecha_limpia = limpiar_fecha_para_calendario(row['Fecha'])
        
        if fecha_limpia:
            # Detectamos si tiene cargado algo en el campo de invitación
            invitacion_val = str(row.get('Invitación a participar', '')).strip()
            tiene_invitacion = invitacion_val != "" and invitacion_val.lower() != "nan"
            
            # Formatear la hora para el título si existe
            hora_val = str(row.get('Hora', '')).strip()
            tiene_hora = hora_val != "" and hora_val.lower() != "nan"
            prefijo_hora = f"{hora_val} hs - " if tiene_hora else ""
            
            # Si tiene invitación, asignamos el Naranja y el emoji de sobre.
            if tiene_invitacion:
                color_evento = COLOR_CON_INVITACION
                titulo_mostrar = f"✉️ {prefijo_hora}[{row.get('Ciudad', 'Sin especificar')}] {row['Actividad']}"
            else:
                color_evento = colores_prioridad.get(str(row['Prioridad']).upper().strip(), "#333333")
                titulo_mostrar = f"{prefijo_hora}[{row.get('Ciudad', 'Sin especificar')}] {row['Actividad']}"
                
            events.append({
                "title": titulo_mostrar,
                "start": fecha_limpia,
                "end": fecha_limpia,
                "color": color_evento,
                "extendedProps": {
                    "fecha_original": str(row['Fecha']),
                    "hora": hora_val if tiene_hora else "No especificada",
                    "ciudad": str(row.get('Ciudad', 'Sin especificar')),
                    "lugar": str(row.get('Lugar', 'Sin especificar')),
                    "explicacion": str(row.get('Explicación breve de la actividad', row.get('Descripción', ''))),
                    "asistencia": int(row.get('Cantidad de personas estimadas', 0)),
                    "organismo": str(row['Organismo/Actor']),
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
                st.markdown(f"**⏰ Hora:** `{props.get('hora')}`")
                st.markdown(f"**📍 Ciudad:** {props.get('ciudad')}")
                st.markdown(f"**🏢 Lugar físico:** {props.get('lugar')}")
                st.markdown(f"**👥 Cantidad de personas estimadas:** `{props.get('asistencia')}`")
            with col_det2:
                st.markdown(f"**🚨 Prioridad:** `{props.get('prioridad')}`")
                st.markdown(f"**⚙️ Estado:** `{props.get('estado')}`")
                st.markdown(f"**✉️ Invitación a participar:** `{props.get('invitacion')}`")
                st.markdown(f"**📝 Explicación breve:** {props.get('explicacion')}")
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
                f_hora = st.time_input("Hora del evento", value=time(9, 0)) 
                f_actividad = st.text_input("Nombre de la Actividad", placeholder="Ej: Lanzamiento Programa Desafíos")
                f_ciudad = st.text_input("Ciudad", placeholder="Ej: General Roca")
                f_lugar = st.text_input("Lugar / Espacio Físico", placeholder="Ej: Aula Magna UNRN")
                f_asistencia = st.number_input("Cantidad de personas estimadas", min_value=0, step=10, value=50)
                
            with col2:
                f_organismo = st.text_input("Organismo / Actor principal", placeholder="Ej: Secretaría de Estado de Energía y Ambiente")
                f_prioridad = st.selectbox("Nivel de Prioridad", ["ALTA", "INTERMEDIA", "BAJA"])
                f_estado = st.selectbox("Estado Actual", ["Pendiente", "En curso", "Finalizado", "Suspendido"])
                f_publico = st.text_input("Público Objetivo", placeholder="Ej: Jóvenes y adultos")
                f_invitacion = st.text_input("Invitación a participar", placeholder="Ej: Gobernador + Secretaria de EEYA")
                f_explicacion = st.text_area("Explicación breve de la actividad")
                
            submitted = st.form_submit_button("💾 Guardar en la Agenda")
            
            if submitted:
                if not f_actividad or not f_ciudad:
                    st.error("Por favor, completa obligatoriamente los campos 'Actividad' y 'Ciudad'.")
                else:
                    hora_formateada = f_hora.strftime("%H:%M")
                    
                    nueva_actividad = {
                        "Fecha": str(f_fecha),
                        "Hora": hora_formateada,
                        "Semana": int(f_fecha.isocalendar()[1]),
                        "Actividad": f_actividad,
                        "Ciudad": f_ciudad.strip(),
                        "Lugar": f_lugar.strip(),
                        "Explicación breve de la actividad": f_explicacion,
                        "Cantidad de personas estimadas": int(f_asistencia),
                        "Organismo/Actor": f_organismo,
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
                ciudad_txt = str(row.get('Ciudad', 'Sin especificar'))
                opciones_actividades.append(f"{idx} | [{ciudad_txt}] {row['Actividad']} ({row['Fecha']})")
                
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
                        
                        # Manejo de la Hora
                        hora_previa_str = str(registro_actual.get('Hora', '09:00')).strip()
                        try:
                            hora_previa = datetime.strptime(hora_previa_str, "%H:%M").time()
                        except:
                            hora_previa = time(9, 0)
                        ed_hora = st.time_input("Hora", value=hora_previa)
                        
                        ed_actividad = st.text_input("Nombre de la Actividad", value=str(registro_actual['Actividad']))
                        ed_ciudad = st.text_input("Ciudad", value=str(registro_actual.get('Ciudad', '')))
                        ed_lugar = st.text_input("Lugar / Espacio Físico", value=str(registro_actual.get('Lugar', '')))
                        
                        try:
                            asistencia_previa = int(registro_actual.get('Cantidad de personas estimadas', 50))
                        except:
                            asistencia_previa = 50
                        ed_asistencia = st.number_input("Cantidad de personas estimadas", min_value=0, value=asistencia_previa)
                        
                    with col2_ed:
                        ed_organismo = st.text_input("Organismo / Actor", value=str(registro_actual['Organismo/Actor']))
                        opciones_prioridad = ["ALTA", "INTERMEDIA", "BAJA"]
                        def_prio = opciones_prioridad.index(str(registro_actual['Prioridad']).upper().strip()) if str(registro_actual['Prioridad']).upper().strip() in opciones_prioridad else 1
                        
                        opciones_estado = ["Pendiente", "En curso", "Finalizado", "Suspendido"]
                        def_est = opciones_estado.index(str(registro_actual['Estado']).capitalize().strip()) if str(registro_actual['Estado']).capitalize().strip() in opciones_estado else 0
                        
                        ed_prioridad = st.selectbox("Prioridad", opciones_prioridad, index=def_prio)
                        ed_estado = st.selectbox("Estado", opciones_estado, index=def_est)
                        ed_publico = st.text_input("Público Destinatario", value=str(registro_actual['Público Destinatario']))
                        ed_invitacion = st.text_input("Invitación a participar", value=str(registro_actual.get('Invitación a participar', '')))
                        ed_explicacion = st.text_area("Explicación breve de la actividad", value=str(registro_actual.get('Explicación breve de la actividad', registro_actual.get('Descripción', ''))))
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        boton_actualizar = st.form_submit_button("🔄 Actualizar Cambios")
                    with col_btn2:
                        boton_eliminar = st.form_submit_button("❌ Eliminar Actividad permanentemente")
                        
                    if boton_actualizar:
                        st.session_state.agenda.at[idx_seleccionado, 'Fecha'] = str(ed_fecha)
                        st.session_state.agenda.at[idx_seleccionado, 'Hora'] = ed_hora.strftime("%H:%M")
                        st.session_state.agenda.at[idx_seleccionado, 'Semana'] = int(ed_fecha.isocalendar()[1])
                        st.session_state.agenda.at[idx_seleccionado, 'Actividad'] = ed_actividad
                        st.session_state.agenda.at[idx_seleccionado, 'Ciudad'] = ed_ciudad.strip()
                        st.session_state.agenda.at[idx_seleccionado, 'Lugar'] = ed_lugar.strip()
                        st.session_state.agenda.at[idx_seleccionado, 'Explicación breve de la actividad'] = ed_explicacion
                        st.session_state.agenda.at[idx_seleccionado, 'Cantidad de personas estimadas'] = int(ed_asistencia)
                        st.session_state.agenda.at[idx_seleccionado, 'Organismo/Actor'] = ed_organismo
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
            ciudades_unicas = sorted(list(df_filtrado['Ciudad'].astype(str).str.strip().unique()))
            list_ciudades = ["Todas"] + [c for c in ciudades_unicas if c != 'nan' and c != '']
            filtro_ciudad = st.selectbox("Filtrar por Ciudad", list_ciudades, key="filter_loc_tab4")
            
        # Filtrar por Ciudad
        if filtro_ciudad != "Todas":
            df_filtrado = df_filtrado[df_filtrado['Ciudad'].str.strip() == filtro_ciudad]
            
        # Filtrar palabra clave
        if search_query:
            df_filtrado = df_filtrado[
                df_filtrado['Actividad'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Explicación breve de la actividad'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Lugar'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Ciudad'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Organismo/Actor'].astype(str).str.lower().str.contains(search_query) |
                df_filtrado['Invitación a participar'].astype(str).str.lower().str.contains(search_query)
            ]
            
        # Renderizado de la tabla
        st.dataframe(df_filtrado, use_container_width=True)
        
        # Botones de descarga de reportes
        col_down1, col_down2, col_down3 = st.columns([1.5, 1.5, 3])
        
        with col_down1:
            # Exportador a Excel seguro
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Agenda')
            excel_bytes = output_excel.getvalue()
            
            st.download_button(
                label="📥 Descargar Excel de Datos",
                data=excel_bytes,
                file_name="agenda_territorial_filtrada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_excel",
                use_container_width=True
            )
            
        with col_down2:
            # Botón de Word condicionado a que la instalación en requirements.txt haya finalizado
            if LIBRERIA_DOCX_LISTA:
                try:
                    word_bytes = crear_reporte_word_areas(df_filtrado)
                    st.download_button(
                        label="📝 Descargar Reporte en Word",
                        data=word_bytes,
                        file_name="Reporte_Agenda_Territorial_UPEU.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_download_word",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error("Error temporal al procesar el archivo de Word.")
            else:
                st.warning("Instalando componente de Word en el servidor. Aguarda unos instantes y recarga la página.")
                
    else:
        st.warning("La base de datos está vacía o cargando. Prueba pulsar el botón '🔄 Sincronizar Excel' arriba a la derecha.")
