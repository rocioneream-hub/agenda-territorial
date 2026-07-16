import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from datetime import datetime, time, date
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

def crear_reporte_word_areas(df, filtrar_desde_hoy=False):
    """Genera un reporte DOCX formal filtrando opcionalmente desde la fecha de hoy en adelante."""
    doc = docx.Document()
    hoy_date = date.today()
    
    # Clonamos el DataFrame para no alterar la visualización en pantalla de la tabla
    df_procesar = df.copy()
    
    # Aplicación del filtro dinámico temporal si se solicita "Desde hoy en adelante"
    if filtrar_desde_hoy:
        fechas_limpias = []
        for idx, row in df_procesar.iterrows():
            f_str = limpiar_fecha_para_calendario(row['Fecha'])
            try:
                fechas_limpias.append(datetime.strptime(f_str, "%Y-%m-%d").date())
            except:
                fechas_limpias.append(date(2000, 1, 1))
        
        df_procesar['_fecha_comparar'] = fechas_limpias
        df_procesar = df_procesar[df_procesar['_fecha_comparar'] >= hoy_date].drop(columns=['_fecha_comparar'])
    
    # Configuración de márgenes estándar
    try:
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
    except:
        pass
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # MEMBRETADO INSTITUCIONAL
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_gob = p_header.add_run("GOBIERNO DE LA PROVINCIA DE RÍO NEGRO\n")
    run_gob.font.bold = True
    run_gob.font.size = Pt(10)
    try:
        run_gob.font.color.rgb = docx.shared.RGBColor(106, 198, 79) # Verde RN (#6AC64F)
    except:
        pass
    
    run_sub = p_header.add_run("Ministerio de Educación y Derechos Humanos\nUnidad Provincial de Enlace con Universidades (UPEU)\n")
    run_sub.font.size = Pt(9.5)
    try:
        run_sub.font.color.rgb = docx.shared.RGBColor(100, 100, 100)
    except:
        pass
    
    # Línea divisoria decorativa
    try:
        p_line = doc.add_paragraph()
        p_line_border = OxmlElement('w:pBdr')
        bottom_border = OxmlElement('w:bottom')
        bottom_border.set(qn('w:val'), 'single')
        bottom_border.set(qn('w:sz'), '8')
        bottom_border.set(qn('w:space'), '1')
        bottom_border.set(qn('w:color'), '007BE0') # Azul RN
        p_line_border.append(bottom_border)
        p_line._p.get_or_add_pPr().append(p_line_border)
    except:
        pass
    
    # TÍTULOS PRINCIPALES
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    titulo_texto = "REPORTE PLANIFICACIÓN TERRITORIAL FUTURA" if filtrar_desde_hoy else "REPORTE COMPLETO DE ACTIVIDADES TERRITORIALES"
    run_title = p_title.add_run(titulo_texto)
    run_title.font.bold = True
    run_title.font.size = Pt(15)
    try:
        run_title.font.color.rgb = docx.shared.RGBColor(0, 123, 224) # Azul RN
    except:
        pass
    p_title.paragraph_format.space_after = Pt(2)
    
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    rango_texto = f"Acciones programadas desde el {hoy_date.strftime('%d/%m/%Y')} en adelante" if filtrar_desde_hoy else "Historial completo de gestión"
    run_date = p_date.add_run(f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')} | {rango_texto}")
    run_date.font.italic = True
    run_date.font.size = Pt(9.5)
    p_date.paragraph_format.space_after = Pt(24)
    
    # CUADRO RESUMEN DE ASISTENCIA GLOBAL
    total_acciones = len(df_procesar)
    try:
        total_personas = df_procesar['Cantidad de personas estimadas'].fillna(0).astype(int).sum()
    except:
        total_personas = 0
    
    table_resumen = doc.add_table(rows=2, cols=2)
    hdr_res = table_resumen.rows[0].cells
    hdr_res[0].text = "Acciones en el Reporte"
    hdr_res[1].text = "Proyección de Asistentes Global"
    
    for cell in hdr_res:
        set_cell_background(cell, "F5F5F5")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(10)
        
    row_res = table_resumen.rows[1].cells
    row_res[0].text = str(total_acciones)
    row_res[1].text = f"{total_personas:,} personas"
    
    for cell in row_res:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(12)
        try:
            p.runs[0].font.color.rgb = docx.shared.RGBColor(0, 123, 224)
        except:
            pass
        
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    # DESGLOSE DE LAS ACTIVIDADES
    doc.add_heading("Fichas de Planificación de Actividades", level=2)
    
    if total_acciones > 0:
        df_ordenado = df_procesar.copy()
        try:
            df_ordenado['Fecha'] = df_ordenado['Fecha'].astype(str).str.strip()
            df_ordenado = df_ordenado.sort_values(by='Fecha').reset_index(drop=True)
        except:
            pass
        
        for idx, row in df_ordenado.iterrows():
            p_act = doc.add_paragraph()
            p_act.paragraph_format.space_before = Pt(14)
            p_act.paragraph_format.space_after = Pt(4)
            
            act_titulo = str(row.get('Actividad', 'Sin Nombre'))
            run_header = p_act.add_run(f"📌 Actividad {idx+1}: {act_titulo}")
            run_header.font.bold = True
            run_header.font.size = Pt(11)
            try:
                run_header.font.color.rgb = docx.shared.RGBColor(0, 123, 224)
            except:
                pass
            
            # Tabla de Ficha Técnica Individual (6 Campos Requeridos)
            ficha_table = doc.add_table(rows=6, cols=2)
            ficha_table.style = 'Light Shading Accent 1'
            
            # Campo 1: Actividad
            ficha_table.rows[0].cells[0].text = "Actividad:"
            ficha_table.rows[0].cells[1].text = act_titulo
            
            # Campo 2: Ciudad
            ficha_table.rows[1].cells
