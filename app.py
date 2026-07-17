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
# 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS (IDENTIDAD VISUAL OFICIAL)
# ==========================================
st.set_page_config(
    layout="wide", 
    page_title="Calendario territorial UPEU", 
    page_icon="🗓️"
)

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
# 3. AUTENTICACIÓN Y CONEXIÓN CON LA API OFICIAL
# ==========================================
def obtener_servicio_sheets():
    alcance = ["https://www.googleapis.com/auth/spreadsheets"]
    credenciales_dict = dict(st.secrets["gcp_service_account"])
    
    # Parche por si StreamlitCloud deforma los saltos de línea al guardar
    if "\\n" in credenciales_dict["private_key"]:
        credenciales_dict["private_key"] = credenciales_dict["private_key"].replace("\\n", "\n")
        
    credenciales = Credentials.from_service_account_info(credenciales_dict, scopes=alcance)
    return build('sheets', 'v4', credentials=credenciales)

def load_data_from_sheets():
    try:
        servicio = obtener_servicio_sheets()
        planilla_id = st.secrets["spreadsheet"]["id"]
        
        # Leemos el rango extendido de la primera pestaña de la hoja de cálculo
        resultado = servicio.spreadsheets().values().get(spreadsheetId=planilla_id, range="A1:Z2000").execute()
        filas = resultado.get('values', [])
        
        if not filas:
            return pd.DataFrame()
            
        encabezados = [str(h).strip() for h in filas[0]]
        datos_crudos = filas[1:]
        
        # Emparejamos filas que tengan celdas vacías al final para evitar descalces en Pandas
        datos_normalizados = []
        for f in datos_crudos:
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
        
        # Limpiamos primero el rango viejo para evitar remanentes si se achica la base de datos
        servicio.spreadsheets().values().clear(spreadsheetId=planilla_id, range="A1:Z2000").execute()
        
        # Sobrescribimos de manera atómica con codificación nativa
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
        meses_dict = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
            "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }
        anio_detectado = date.today().year
        match_anio = re.search(r"\b(202\d)\b", val_str)
        if match_anio: anio_detectado = int(match_anio.group(1))
        for mes_nombre, mes_num in meses_dict.items():
            if mes_nombre in val_str.lower(): return f"{anio_detectado}-{str(mes_num).zfill(2)}-01"
        return f"{anio_detectado}-01-01"

    val_str = re.sub(r"\s*/\s*", "/", val_str)
    match_normal = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", val_str)
    if match_normal: return f"{match_normal.group(3)}-{match_normal.group(2).zfill(2)}-{match_normal.group(1).zfill(2)}"
    return None

if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data_from_sheets()

# ==========================================
# EXPORTACIÓN REPORTE WORD COMPLETO
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
    run_gob = p_header.add_run("GOBIERNO DE LA PROVINCIA DE RÍO NEGRO\n")
    run_gob.font.bold = True; run_gob.font.size = Pt(10)
    p_header.add_run("Ministerio de Educación y Derechos Humanos\nUnidad Provincial de Enlace con Universidades (UPEU)\n")
    
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(titulo_personalizado.upper())
    run_title.font.bold = True; run_title.font.size = Pt(15)
    
    p_date = doc.add_paragraph()
    p_date.add_run(f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')} | {aclaracion_rango}").font.italic = True
    
    total_acciones = len(df)
    table_resumen = doc.add_table(rows=2, cols=2)
    table_resumen.rows[0].cells[0].text = "Acciones en el Reporte"
    table_resumen.rows[0].cells[1].text = "Proyección de Asistentes Global"
    set_cell_background(table_resumen.rows[0].cells[0], "F5F5F5")
    set_cell_background(table_resumen.rows[0].cells[1], "F5F5F5")
    
    table_resumen.rows[1].cells[0].text = str(total_acciones)
    try: total_personas = df['Cantidad de personas estimadas'].fillna(0).astype(int).sum()
    except: total_personas = 0
    table_resumen.rows[1].cells[1].text = f"{total_personas:,} personas"
    
    doc.add_paragraph()
    doc.add_heading("Fichas de Planificación de Actividades", level=2)
    
    if total_acciones > 0:
        df_ordenado = df.copy()
        try:
            df_ordenado['Fecha_Limpia'] = df_ordenado['Fecha'].apply(limpiar_fecha_para_calendario)
            df_ordenado = df_ordenado.sort_values(by='Fecha_Limpia').reset_index(drop=True)
        except: pass
        
        for idx, row in df_ordenado.iterrows():
            p_act = doc.add_paragraph()
            run_h = p_act.add_run(f"📌 Actividad {idx+1}: {row.get('Actividad', 'Sin Nombre')}")
            run_h.font.bold = True
            
            ficha_table = doc.add_table(rows=6, cols=2)
            ficha_table.style = 'Light Shading Accent 1'
            ficha_table.rows[0].cells[0].text = "Actividad:"; ficha_table.rows[0].cells[1].text = str(row.get('Actividad', ''))
            ficha_table.rows[1].cells[0].text = "Ciudad:"; ficha_table.rows[1].cells[1].text = str(row.get('Ciudad', ''))
            
            fecha_val = str(row.get('Fecha', 'Sin especificar')).strip().split(" ")[0]
            fecha_mostrar = "Sin especificar (A coordinar)" if ("sin especificar" in fecha_val.lower() or fecha_val == "") else fecha_val
            hora_val = str(row.get('Hora', 'Sin especificar')).strip()
            hora_text = " - Sin especificar" if ("sin especificar" in hora_val.lower() or hora_val == "") else f" - {hora_val} hs"
            
            ficha_table.rows[2].cells[0].text = "Fecha:"; ficha_table.rows[2].cells[1].text = f"{fecha_mostrar}{hora_text}"
            ficha_table.rows[3].cells[0].text = "Lugar:"; ficha_table.rows[3].cells[1].text = str(row.get('Lugar', ''))
            ficha_table.rows[4].cells[0].text = "Explicación breve de la actividad:"; ficha_table.rows[4].cells[1].text = str(row.get('Explicación breve de la actividad', ''))
            ficha_table.rows[5].cells[0].text = "Cantidad de personas estimadas:"; ficha_table.rows[5].cells[1].text = f"{int(row.get('Cantidad de personas estimadas', 0)):,} asistentes"
            
            for r in ficha_table.rows: r.cells[0].paragraphs[0].runs[0].font.bold = True
            doc.add_paragraph()
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# REPORTE DE WHATSAPP COMPLETO
# ==========================================
def generar_mensaje_whatsapp(df, titulo_cabecera="Agenda completa de actividades"):
    df_procesar = df.copy()
    try:
        df_procesar['Fecha_Limpia'] = df_procesar['Fecha'].apply(limpiar_fecha_para_calendario)
        df_procesar = df_procesar.sort_values(by='Fecha_Limpia').reset_index(drop=True)
    except: pass
        
    lines = ["🏛️ *UPEU - PLANIFICACIÓN TERRITORIAL*", f"📅 *{titulo_cabecera}*:", "─────────────────────"]
    total_eventos = len(df_procesar)
    if total_eventos == 0:
        lines.append("*(Sin actividades planificadas)*")
    else:
        for idx, row in df_procesar.iterrows():
            act_titulo = str(row.get('Actividad', 'Sin Nombre')).upper().strip()
            fecha_val = str(row.get('Fecha', 'Sin especificar')).strip()
            fecha_mostrar = "Sin especificar (A coordinar por el territorio)" if ("sin especificar" in fecha_val.lower() or fecha_val == "") else fecha_val
            hora_val = str(row.get('Hora', 'Sin especificar')).strip()
            hora_txt = " - Horario sin especificar" if ("sin especificar" in hora_val.lower() or hora_val == "") else f" - {hora_val} hs"
            
            lines.append(f"📌 *Actividad {idx+1}:* {act_titulo}")
            lines.append(f"📍 *Ciudad:* {str(row.get('Ciudad', 'Sin especificar')).strip()}")
            lines.append(f"📅 *Fecha:* {fecha_mostrar}{hora_txt}")
            lines.append(f"🏢 *Lugar:* {str(row.get('Lugar', 'Sin especificar')).strip()}")
            lines.append(f"📝 *Explicación breve de la actividad:* {str(row.get('Explicación breve de la actividad', '')).strip()}")
            lines.append(f"👥 *Cantidad de personas estimadas:* {int(row.get('Cantidad de personas estimadas', 0)):,} personas")
            
            if idx < total_eventos - 1: lines.append("\n🔸  🔸  🔸  🔸  🔸\n")
    lines.append("─────────────────────")
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
    st.markdown("<span class='hashtag-gestion'>#gobiernodelosrionegrinos</span>", unsafe_allow_html=True) 

with col_title_right:
    st.write(""); st.write("")
    if st.button("🔄 Sincronizar Google Sheets", use_container_width=True):
        st.session_state.agenda = load_data_from_sheets()
        st.success("¡Base de datos en la nube sincronizada!")
        st.rerun()

if es_editor:
    tab1, tab2, tab3, tab4 = st.tabs(["🗓️ Vista de Calendario", "✍️ Carga Rápida de Actividad", "✏️ Modificar / Eliminar Actividad", "📊 Base de Datos Completa"])
else:
    tab1, tab4 = st.tabs(["🗓️ Vista de Calendario", "📊 Base de Datos Completa"])
    tab2, tab3 = None, None

# ------------------------------------------
# TAB 1: CALENDARIO INTERACTIVO RESTAURADO
# ------------------------------------------
with tab1:
    st.header("Planificación Territorial")
    events = []
    colores_prioridad = {"ALTA": "#007BE0", "INTERMEDIA": "#333333", "BAJA": "#6AC64F"}
    
    if len(st.session_state.agenda) > 0:
        for idx, row in st.session_state.agenda.iterrows():
            fecha_limpia = limpiar_fecha_para_calendario(row['Fecha'])
            if fecha_limpia:
                invitacion_val = str(row.get('Invitación a participar', '')).strip()
                tiene_invitacion = invitacion_val != "" and invitacion_val.lower() != "nan"
                
                act_txt = str(row['Actividad'])
                if "sin especificar" in str(row['Fecha']).lower(): act_txt = f"📍 [FECHA FLEXIBLE] {act_txt}"
                    
                color_evento = "#FF7A00" if tiene_invitacion else colores_prioridad.get(str(row['Prioridad']).upper().strip(), "#333333")
                events.append({
                    "title": f"[{row.get('Ciudad', 'Sin especificar')}] {act_txt}",
                    "start": fecha_limpia, "end": fecha_limpia, "color": color_evento,
                    "extendedProps": {
                        "fecha_original": str(row['Fecha']), "hora": str(row['Hora']),
                        "ciudad": str(row['Ciudad']), "lugar": str(row['Lugar']),
                        "explicacion": str(row['Explicación breve de la actividad']),
                        "asistencia": int(row['Cantidad de personas estimadas']), "invitacion": invitacion_val if tiene_invitacion else "Ninguna"
                    }
                })

    if len(events) > 0:
        state = calendar(events=events, options={"initialView": "dayGridMonth", "locale": "es"}, key="calendar_agenda")
        if state.get("eventClick"):
            props = state["eventClick"]["event"]["extendedProps"]
            st.markdown("---")
            st.subheader("🔍 Detalle de la Actividad Seleccionada")
            st.write(f"**📌 Actividad:** {state['eventClick']['event']['title']}")
            st.write(f"**📅 Fecha original:** `{props.get('fecha_original')}` | **⏰ Hora:** `{props.get('hora')}`")
            st.write(f"**🏢 Lugar:** {props.get('lugar')} ({props.get('ciudad')})")
            st.write(f"**📝 Explicación:** {props.get('explicacion')}")
            st.write(f"**✉️ Invitación:** `{props.get('invitacion')}`")
    else:
        st.warning("No hay actividades con fechas válidas para renderizar el calendario.")

# ------------------------------------------
# TAB 2: CARGA RÁPIDA DETALLADA RESTAURADA
# ------------------------------------------
if es_editor and tab2 is not None:
    with tab2:
        st.header("Registrar Nueva Actividad")
        with st.form("nuevo_evento_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**📅 Fecha del Evento**")
                fecha_sin_especificar = st.checkbox("Dejar fecha sin especificar (A coordinar)", value=False)
                mes_propuesto = st.selectbox("Mes de referencia:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], index=datetime.today().month-1)
                anio_propuesto = st.selectbox("Año:", [2026, 2027])
                f_fecha = st.date_input("Elegir fecha fija", datetime.today())
                
                st.markdown("**⏰ Horario**")
                hora_sin_especificar = st.checkbox("Dejar hora sin especificar (A coordinar)", value=False)
                f_hora = st.time_input("Hora fija", value=time(9, 0)) 
                
                f_actividad = st.text_input("Nombre de la Actividad")
                f_ciudad = st.text_input("Ciudad")
                f_lugar = st.text_input("Lugar / Espacio Físico")
                f_asistencia = st.number_input("Cantidad de personas estimadas", min_value=0, value=50)
            with col2:
                f_organismo = st.text_input("Organismo / Actor")
                f_prioridad = st.selectbox("Prioridad", ["INTERMEDIA", "ALTA", "BAJA"])
                f_estado = st.selectbox("Estado", ["Pendiente", "En curso", "Finalizado", "Suspendido"])
                f_publico = st.text_input("Público Objetivo")
                f_invitacion = st.text_input("Invitación a participar")
                f_explicacion = st.text_area("Explicación breve de la actividad")
                
            submitted = st.form_submit_button("💾 Guardar en la Nube de Google Sheets")
            if submitted:
                if not f_actividad or not f_ciudad:
                    st.error("Completa los campos obligatorios: 'Actividad' y 'Ciudad'.")
                else:
                    if fecha_sin_especificar:
                        meses_dict = {"Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12}
                        fecha_guardar = f"Sin especificar ({mes_propuesto} {anio_propuesto})"
                        fecha_tecnica = date(anio_propuesto, meses_dict[mes_propuesto], 1)
                    else:
                        fecha_guardar = str(f_fecha); fecha_tecnica = f_fecha
                        
                    hora_guardar = "Sin especificar" if hora_sin_especificar else f_hora.strftime("%H:%M")
                    
                    nueva_fila = pd.DataFrame([{
                        "Fecha": fecha_guardar, "Hora": hora_guardar, "Semana": int(fecha_tecnica.isocalendar()[1]),
                        "Actividad": f_actividad, "Ciudad": f_ciudad.strip(), "Lugar": f_lugar.strip(),
                        "Explicación breve de la actividad": f_explicacion, "Cantidad de personas estimadas": int(f_asistencia),
                        "Organismo/Actor": f_organismo, "Estado": f_estado, "Público Destinatario": f_publico,
                        "Prioridad": f_prioridad, "Invitación a participar": f_invitacion
                    }])
                    
                    df_actualizado = pd.concat([st.session_state.agenda, nueva_fila], ignore_index=True)
                    if guardar_todo_en_sheets(df_actualizado):
                        st.session_state.agenda = load_data_from_sheets()
                        st.success("¡Actividad registrada con éxito mediante API oficial!")
                        st.rerun()

# ------------------------------------------
# TAB 3: MODIFICAR / ELIMINAR REGISTROS RESTAURADO
# ------------------------------------------
if es_editor and tab3 is not None:
    with tab3:
        st.header("Editar / Cancelar Actividades")
        df_sheet = st.session_state.agenda.copy()
        
        if len(df_sheet) > 0:
            opciones_actividades = [f"{idx} | [{row['Ciudad']}] {row['Actividad']} ({row['Fecha']})" for idx, row in df_sheet.iterrows()]
            actividad_seleccionada = st.selectbox("Seleccionar Actividad a Gestionar", opciones_actividades)
            
            if actividad_seleccionada:
                idx_seleccionado = int(actividad_seleccionada.split(" | ")[0])
                registro_actual = df_sheet.loc[idx_seleccionado]
                
                with st.form("form_edicion"):
                    col1_ed, col2_ed = st.columns(2)
                    with col1_ed:
                        ed_fecha_sin = st.checkbox("Dejar fecha sin especificar", value=("sin especificar" in str(registro_actual['Fecha']).lower()))
                        ed_actividad = st.text_input("Actividad", value=str(registro_actual['Actividad']))
                        ed_ciudad = st.text_input("Ciudad", value=str(registro_actual['Ciudad']))
                        ed_lugar = st.text_input("Lugar", value=str(registro_actual['Lugar']))
                        ed_asistencia = st.number_input("Asistentes", value=int(registro_actual['Cantidad de personas estimadas']))
                    with col2_ed:
                        ed_organismo = st.text_input("Organismo", value=str(registro_actual['Organismo/Actor']))
                        ed_prioridad = st.selectbox("Prioridad", ["INTERMEDIA", "ALTA", "BAJA"], index=["INTERMEDIA", "ALTA", "BAJA"].index(str(registro_actual['Prioridad']).upper().strip() if str(registro_actual['Prioridad']) else "INTERMEDIA"))
                        ed_estado = st.selectbox("Estado", ["Pendiente", "En curso", "Finalizado", "Suspendido"])
                        ed_invitacion = st.text_input("Invitación a participar", value=str(registro_actual['Invitación a participar']))
                        ed_explicacion = st.text_area("Explicación breve", value=str(registro_actual['Explicación breve de la actividad']))
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1: boton_actualizar = st.form_submit_button("🔄 Actualizar en Google Sheets")
                    with col_btn2: boton_eliminar = st.form_submit_button("❌ Eliminar Permanentemente")
                        
                    if boton_actualizar:
                        df_sheet.at[idx_seleccionado, 'Actividad'] = ed_actividad
                        df_sheet.at[idx_seleccionado, 'Ciudad'] = ed_ciudad
                        df_sheet.at[idx_seleccionado, 'Lugar'] = ed_lugar
                        df_sheet.at[idx_seleccionado, 'Cantidad de personas estimadas'] = ed_asistencia
                        df_sheet.at[idx_seleccionado, 'Organismo/Actor'] = ed_organismo
                        df_sheet.at[idx_seleccionado, 'Prioridad'] = ed_prioridad
                        df_sheet.at[idx_seleccionado, 'Estado'] = ed_estado
                        df_sheet.at[idx_seleccionado, 'Invitación a participar'] = ed_invitacion
                        df_sheet.at[idx_seleccionado, 'Explicación breve de la actividad'] = ed_explicacion
                        if ed_fecha_sin: df_sheet.at[idx_seleccionado, 'Fecha'] = "Sin especificar (A coordinar)"
                        
                        if guardar_todo_en_sheets(df_sheet):
                            st.session_state.agenda = load_data_from_sheets()
                            st.success("¡Registro modificado en la nube con codificación nativa!")
                            st.rerun()
                        
                    if boton_eliminar:
                        df_sheet = df_sheet.drop(idx_seleccionado).reset_index(drop=True)
                        if guardar_todo_en_sheets(df_sheet):
                            st.session_state.agenda = load_data_from_sheets()
                            st.warning("¡Registro eliminado de la planilla de Google!")
                            st.rerun()

# ------------------------------------------
# TAB 4: BUSCADOR, MOTORES DE FILTRADO Y REPORTES RESTAURADOS
# ------------------------------------------
with tab4:
    st.header("Buscador y Reportes")
    df_filtrado = st.session_state.agenda.copy()
    
    if len(df_filtrado) > 0:
        search_query = st.text_input("Buscar por palabra clave...").lower()
        if search_query: df_filtrado = df_filtrado[df_filtrado['Actividad'].astype(str).str.lower().str.contains(search_query)]
            
        st.dataframe(df_filtrado, use_container_width=True)
        st.markdown("### 📤 Generar y Exportar Documentos")
        
        try: semanas_disponibles = sorted([int(s) for s in df_filtrado['Semana'].dropna().unique() if str(s).strip() != "" and str(s).lower() != "nan"])
        except: semanas_disponibles = []
            
        col_config, col_semana_selector, col_down1, col_down2 = st.columns([2.2, 1.5, 1.3, 1.3])
        with col_config:
            rango_reporte = st.radio("Alcance temporal:", ["Agenda Completa (Historial + Futuro)", "Desde Hoy hacia adelante", "Filtrar por una semana específica"])
            solo_con_invitacion = st.checkbox("🔍 Filtrar SOLO actividades con 'Invitación a participar'")
            
        df_descarga = df_filtrado.copy()
        titulo_word, titulo_whatsapp, rango_aclaracion_word = "REPORTE PLANIFICACION TERRITORIAL", "Cronograma de actividades", "Coordinación Interinstitucional"
        
        if rango_reporte == "Desde Hoy hacia adelante":
            hoy_date = date.today()
            fechas_comp = [datetime.strptime(limpiar_fecha_para_calendario(r['Fecha']), "%Y-%m-%d").date() if limpiar_fecha_para_calendario(r['Fecha']) else date(2000,1,1) for idx, r in df_descarga.iterrows()]
            df_descarga['_comparar'] = fechas_comp
            df_descarga = df_descarga[df_descarga['_comparar'] >= hoy_date].drop(columns=['_comparar'])
            titulo_word, titulo_whatsapp = "REPORTE PLANIFICACION TERRITORIAL", "Cronograma de actividades"
        elif rango_reporte == "Agenda Completa (Historial + Futuro)":
            titulo_word, titulo_whatsapp = "REPORTE COMPLETO DE ACTIVIDADES TERRITORIALES", "Agenda completa de actividades"
        elif rango_reporte == "Filtrar por una semana específica" and len(semanas_disponibles) > 0:
            with col_semana_selector:
                semana_elegida = st.selectbox("Seleccionar Semana:", semanas_disponibles, format_func=lambda x: f"Semana {x}")
                df_descarga = df_descarga[df_descarga['Semana'] == semana_elegida]
                titulo_word, titulo_whatsapp = f"REPORTE PLANIFICACION TERRITORIAL - SEMANA {semana_elegida}", f"Planificación Territorial - Semana {semana_elegida}"
                
        if solo_con_invitacion:
            df_descarga = df_descarga[df_descarga['Invitación a participar'].notna() & (df_descarga['Invitación a participar'].astype(str).str.strip() != "")]
            titulo_word += " - PROTOCOLO"; titulo_whatsapp += " (Protocolo)"
            
        with col_down1:
            st.write(""); st.write("")
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer: df_descarga.to_excel(writer, index=False, sheet_name='Agenda')
            st.download_button(label="📥 Descargar Excel", data=output_excel.getvalue(), file_name="agenda_territorial.xlsx", use_container_width=True)
            
        with col_down2:
            st.write(""); st.write("")
            if LIBRERIA_DOCX_LISTA:
                word_bytes = crear_reporte_word_areas(df_descarga, titulo_personalizado=titulo_word, aclaracion_rango=rango_aclaracion_word)
                st.download_button(label="📝 Descargar Reporte Word", data=word_bytes, file_name="Reporte_Planificacion.docx", use_container_width=True)
                
        st.markdown("---")
        st.markdown("### 💬 Copiar Reporte para WhatsApp")
        mensaje_whatsapp_generado = generar_mensaje_whatsapp(df_descarga, titulo_cabecera=titulo_whatsapp)
        st.code(mensaje_whatsapp_generado, language="text")
