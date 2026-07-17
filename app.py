import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
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
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(layout="wide", page_title="Calendario territorial UPEU", page_icon="🗓️")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Figtree:ital,wght@0,300..900;1,300..900&display=swap');
    html, body, [class*="css"], .stMarkdown, p, div { font-family: 'Figtree', sans-serif !important; }
    .stApp, .main, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stMainBlockContainer"] { background-color: #E8E8E8 !important; background: #E8E8E8 !important; }
    div[data-testid="stForm"] { background-color: #FFFFFF !important; border-radius: 8px !important; padding: 30px !important; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.05) !important; border: 1px solid #D1D5DB !important; }
    h1 { color: #000000 !important; font-weight: 800 !important; font-size: 2.2rem !important; }
    h2, h3 { color: #007BE0 !important; font-weight: 700 !important; }
    button[kind="primary"] { background-color: #6AC64F !important; color: #FFFFFF !important; font-weight: 700 !important; border-radius: 6px !important; border: none !important; }
    .hashtag-gestion { color: #6AC64F !important; font-weight: 800; font-size: 1.1rem; }
    div[data-baseweb="tab-list"] button[aria-selected="true"] { color: #007BE0 !important; border-bottom-color: #007BE0 !important; }
    div[data-baseweb="tab-list"] button[aria-selected="false"] { color: #333333 !important; }
    .fc .fc-button, .fc .fc-button-primary, button.fc-button { background-color: #007BE0 !important; border-color: #007BE0 !important; color: #FFFFFF !important; font-family: 'Figtree', sans-serif !important; font-weight: 600 !important; }
    .fc .fc-button:hover, button.fc-button:hover { background-color: #6AC64F !important; border-color: #6AC64F !important; }
    .fc .fc-daygrid-day.fc-day-today { background-color: rgba(0, 123, 224, 0.08) !important; }
    </style>
""", unsafe_allow_html=True)

LOGO_FILE = "isologo_RN.svg"

# ==========================================
# 2. CONTROL DE ACCESO
# ==========================================
st.sidebar.header("🔑 Control de Acceso")
password = st.sidebar.text_input("Contraseña de Editor", type="password")
es_editor = (password == "UPEU2026")

# ==========================================
# 3. CONEXIÓN API OFICIAL DE GOOGLE SHEETS
# ==========================================
def obtener_servicio_sheets():
    alcance = ["https://www.googleapis.com/auth/spreadsheets"]
    credenciales_dict = dict(st.secrets["gcp_service_account"])
    
    # Reemplazo seguro por si Streamlit Cloud duplica las barras
    if "\\n" in credenciales_dict["private_key"]:
        credenciales_dict["private_key"] = credenciales_dict["private_key"].replace("\\n", "\n")
        
    credenciales = Credentials.from_service_account_info(credenciales_dict, scopes=alcance)
    return build('sheets', 'v4', credentials=credenciales)

def load_data_from_sheets():
    try:
        servicio = obtener_servicio_sheets()
        planilla_id = st.secrets["spreadsheet"]["id"]
        
        # Llamada nativa a la API para traer la primera pestaña completa
        resultado = servicio.spreadsheets().values().get(spreadsheetId=planilla_id, range="A1:Z1000").execute()
        filas = resultado.get('values', [])
        
        if not filas:
            return pd.DataFrame()
            
        encabezados = [str(h).strip() for h in filas[0]]
        datos = filas[1:]
        
        # Rellenar filas cortas para emparejar el DataFrame
        datos_normalizados = []
        for f in datos:
            fila_extendida = f + [""] * (len(encabezados) - len(f))
            datos_normalizados.append(fila_extendida[:len(encabezados)])
            
        df = pd.DataFrame(datos_normalizados, columns=encabezados)
        
        df['_temp_fecha'] = df['Fecha'].astype(str).str.strip().replace('', None)
        df = df[(~df['_temp_fecha'].isna()) & (df['_temp_fecha'] != 'nan')].drop(columns=['_temp_fecha'])
        
        columnas_requeridas = [
            'Fecha', 'Hora', 'Semana', 'Actividad', 'Ciudad', 'Lugar',
            'Explicación breve de la actividad', 'Cantidad de personas estimadas',
            'Organismo/Actor', 'Estado', 'Público Destinatario', 'Prioridad', 'Invitación a participar'
        ]
        for col in columnas_requeridas:
            if col not in df.columns: df[col] = ""
            
        df['Cantidad de personas estimadas'] = pd.to_numeric(df['Cantidad de personas estimadas'], errors='coerce').fillna(0).astype(int)
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"Error crítico en la lectura de base de datos en la nube: {e}")
        return pd.DataFrame()

def guardar_todo_en_sheets(df):
    try:
        servicio = obtener_servicio_sheets()
        planilla_id = st.secrets["spreadsheet"]["id"]
        
        lista_datos = [df.columns.values.tolist()] + df.astype(str).values.tolist()
        cuerpo = {'values': lista_datos}
        
        # Sobrescribimos el rango de forma atómica y segura
        servicio.spreadsheets().values().update(
            spreadsheetId=planilla_id, range="A1",
            valueInputOption="USER_ENTERED", body=cuerpo
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error al escribir en la nube: {e}")
        return False

def limpiar_fecha_para_calendario(val):
    val_str = str(val).strip()
    if not val_str or val_str == "nan" or val_str == "": return None
    match_iso = re.match(r"^(\d{4}-\d{2}-\d{2})", val_str)
    if match_iso: return match_iso.group(1)
    if "sin especificar" in val_str.lower():
        meses_dict = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
        anio_detectado = date.today().year
        match_anio = re.search(r"\b(202\d)\b", val_str)
        if match_anio: anio_detectado = int(match_anio.group(1))
        for m_n, m_num in meses_dict.items():
            if m_n in val_str.lower(): return f"{anio_detectado}-{str(m_num).zfill(2)}-01"
        return f"{anio_detectado}-01-01"
    val_str = re.sub(r"\s*/\s*", "/", val_str)
    match_normal = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", val_str)
    if match_normal: return f"{match_normal.group(3)}-{match_normal.group(2).zfill(2)}-{match_normal.group(1).zfill(2)}"
    return None

if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data_from_sheets()

# ==========================================
# EXPORTACIÓN REPORTE WORD
# ==========================================
def set_cell_background(cell, color_hex):
    try: cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'))
    except: pass

def crear_reporte_word_areas(df, titulo_personalizado="REPORTE PLANIFICACION TERRITORIAL", aclaracion_rango=""):
    doc = docx.Document()
    try:
        for section in doc.sections:
            section.top_margin = Inches(1); section.bottom_margin = Inches(1)
            section.left_margin = Inches(1); section.right_margin = Inches(1)
    except: pass
    style = doc.styles['Normal']; style.font.name = 'Calibri'; style.font.size = Pt(11)
    p_header = doc.add_paragraph()
    p_header.add_run("GOBIERNO DE LA PROVINCIA DE RÍO NEGRO\n").font.bold = True
    p_header.add_run("Ministerio de Educación y Derechos Humanos\nUnidad Provincial de Enlace con Universidades (UPEU)\n")
    p_title = doc.add_paragraph()
    p_title.add_run(titulo_personalizado.upper()).font.bold = True
    doc.add_heading("Fichas de Planificación de Actividades", level=2)
    
    for idx, row in df.iterrows():
        p_act = doc.add_paragraph()
        p_act.add_run(f"📌 Actividad {idx+1}: {row.get('Actividad', 'Sin Nombre')}").font.bold = True
        ficha_table = doc.add_table(rows=5, cols=2)
        ficha_table.style = 'Light Shading Accent 1'
        ficha_table.rows[0].cells[0].text = "Actividad:"; ficha_table.rows[0].cells[1].text = str(row.get('Actividad', ''))
        ficha_table.rows[1].cells[0].text = "Ciudad:"; ficha_table.rows[1].cells[1].text = str(row.get('Ciudad', ''))
        ficha_table.rows[2].cells[0].text = "Lugar:"; ficha_table.rows[2].cells[1].text = str(row.get('Lugar', ''))
        ficha_table.rows[3].cells[0].text = "Explicación:"; ficha_table.rows[3].cells[1].text = str(row.get('Explicación breve de la actividad', ''))
        ficha_table.rows[4].cells[0].text = "Asistentes:"; ficha_table.rows[4].cells[1].text = f"{int(row.get('Cantidad de personas estimadas', 0)):,} personas"
        doc.add_paragraph()
    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer.getvalue()

def generar_mensaje_whatsapp(df, titulo_cabecera="Agenda"):
    lines = ["🏛️ *UPEU - PLANIFICACIÓN TERRITORIAL*", f"📅 *{titulo_cabecera}*:", "─────────────────────"]
    for idx, row in df.iterrows():
        lines.append(f"📌 *Actividad:* {str(row.get('Actividad', '')).upper()}")
        lines.append(f"📍 *Ciudad:* {str(row.get('Ciudad', ''))} | 🏢 *Lugar:* {str(row.get('Lugar', ''))}")
        lines.append(f"📅 *Fecha:* {str(row.get('Fecha', ''))} - {str(row.get('Hora', ''))} hs")
        if idx < len(df) - 1: lines.append("\n🔸  🔸  🔸  🔸  🔸\n")
    return "\n".join(lines)

# ==========================================
# INTERFAZ GRÁFICA DE USUARIO
# ==========================================
if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=180)
st.markdown("---")

col_title_left, col_title_right = st.columns([4, 1.5])
with col_title_left:
    st.title("Agenda de Planificación Territorial")
    st.markdown("**Unidad Provincial de Enlace con Universidades (UPEU)** | Gobierno de Río Negro")

with col_title_right:
    if st.button("🔄 Sincronizar Google Sheets", use_container_width=True):
        st.session_state.agenda = load_data_from_sheets()
        st.rerun()

if es_editor:
    tab1, tab2, tab3, tab4 = st.tabs(["🗓️ Vista de Calendario", "✍️ Carga Rápida de Actividad", "✏️ Modificar / Eliminar Actividad", "📊 Base de Datos Completa"])
else:
    tab1, tab4 = st.tabs(["🗓️ Vista de Calendario", "📊 Base de Datos Completa"])
    tab2, tab3 = None, None

# TAB 1: CALENDARIO
with tab1:
    events = []
    if len(st.session_state.agenda) > 0:
        for idx, row in st.session_state.agenda.iterrows():
            fecha_limpia = limpiar_fecha_para_calendario(row['Fecha'])
            if fecha_limpia:
                events.append({
                    "title": f"[{row.get('Ciudad', '')}] {row.get('Actividad', '')}",
                    "start": fecha_limpia, "end": fecha_limpia, "color": "#007BE0"
                })
    if events:
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "es"}, key="calendar_upeu")

# TAB 2: CARGA RÁPIDA
if es_editor and tab2:
    with tab2:
        with st.form("nuevo_evento"):
            f_fecha = st.date_input("Fecha", datetime.today())
            f_hora = st.time_input("Hora", time(9,0))
            f_act = st.text_input("Actividad")
            f_ciudad = st.text_input("Ciudad")
            f_lugar = st.text_input("Lugar")
            f_asist = st.number_input("Asistentes", min_value=0, value=10)
            f_expl = st.text_area("Explicación")
            
            if st.form_submit_button("💾 Guardar"):
                nueva_fila = pd.DataFrame([{
                    "Fecha": str(f_fecha), "Hora": f_hora.strftime("%H:%M"), "Semana": int(f_fecha.isocalendar()[1]),
                    "Actividad": f_act, "Ciudad": f_ciudad, "Lugar": f_lugar, "Explicación breve de la actividad": f_expl,
                    "Cantidad de personas estimadas": int(f_asist)
                }])
                df_act = pd.concat([st.session_state.agenda, nueva_fila], ignore_index=True)
                if guardar_todo_en_sheets(df_act):
                    st.session_state.agenda = load_data_from_sheets()
                    st.rerun()

# TAB 3: MODIFICAR
if es_editor and tab3:
    with tab3:
        df_sheet = st.session_state.agenda.copy()
        if len(df_sheet) > 0:
            opciones = [f"{i} | {r['Actividad']}" for i, r in df_sheet.iterrows()]
            sel = st.selectbox("Seleccionar", opciones)
            idx = int(sel.split(" | ")[0])
            with st.form("edit_form"):
                ed_act = st.text_input("Actividad", value=df_sheet.loc[idx, 'Actividad'])
                ed_ciud = st.text_input("Ciudad", value=df_sheet.loc[idx, 'Ciudad'])
                if st.form_submit_button("Actualizar"):
                    df_sheet.at[idx, 'Actividad'] = ed_act
                    df_sheet.at[idx, 'Ciudad'] = ed_ciud
                    if guardar_todo_en_sheets(df_sheet):
                        st.session_state.agenda = load_data_from_sheets()
                        st.rerun()

# TAB 4: DATA
with tab4:
    st.dataframe(st.session_state.agenda, use_container_width=True)
