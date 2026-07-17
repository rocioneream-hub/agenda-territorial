import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, time, date
import re
import os
import io

# Importación de docx
try:
    import docx
    from docx.shared import Pt, Inches
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    LIBRERIA_DOCX_LISTA = True
except ImportError:
    LIBRERIA_DOCX_LISTA = False

# --- CONFIGURACIÓN E IDENTIDAD ---
st.set_page_config(layout="wide", page_title="Calendario territorial UPEU", page_icon="🗓️")
st.markdown("""<style>.stApp { background-color: #E8E8E8 !important; } h1 { color: #000000; } h2 { color: #007BE0; }</style>""", unsafe_allow_html=True)

# --- CONEXIÓN BLINDADA ---
def obtener_servicio_sheets():
    alcance = ["https://www.googleapis.com/auth/spreadsheets"]
    credenciales_dict = dict(st.secrets["gcp_service_account"])
    if "\\n" in credenciales_dict["private_key"]:
        credenciales_dict["private_key"] = credenciales_dict["private_key"].replace("\\n", "\n")
    credenciales = Credentials.from_service_account_info(credenciales_dict, scopes=alcance)
    return build('sheets', 'v4', credentials=credenciales)

def load_data_from_sheets():
    try:
        servicio = obtener_servicio_sheets()
        planilla_id = st.secrets["spreadsheet"]["id"]
        resultado = servicio.spreadsheets().values().get(spreadsheetId=planilla_id, range="A1:Z2000").execute()
        filas = resultado.get('values', [])
        if not filas: return pd.DataFrame()
        df = pd.DataFrame(filas[1:], columns=[str(h).strip() for h in filas[0]])
        # Aseguramos columnas numéricas
        df['Cantidad de personas estimadas'] = pd.to_numeric(df['Cantidad de personas estimadas'], errors='coerce').fillna(0).astype(int)
        return df
    except Exception as e:
        st.error(f"Error en carga: {e}")
        return pd.DataFrame()

# --- FUNCIONES DE LÓGICA ---
def limpiar_fecha(val):
    val_str = str(val).strip()
    match = re.match(r"(\d{4}-\d{2}-\d{2})", val_str)
    return match.group(1) if match else None

if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data_from_sheets()

# --- INTERFAZ ---
st.title("Agenda de Planificación Territorial")
st.markdown("**Unidad Provincial de Enlace con Universidades (UPEU)**")

# Control de acceso
password = st.sidebar.text_input("Contraseña", type="password")
es_editor = (password == "UPEU2026")

tab1, tab4 = st.tabs(["🗓️ Calendario", "📊 Base de Datos"])

with tab1:
    events = []
    for idx, row in st.session_state.agenda.iterrows():
        f = limpiar_fecha(row['Fecha'])
        if f:
            events.append({"title": row['Actividad'], "start": f, "allDay": True})
    calendar(events=events, options={"initialView": "dayGridMonth"})

with tab4:
    st.dataframe(st.session_state.agenda)
    if st.button("🔄 Sincronizar"):
        st.session_state.agenda = load_data_from_sheets()
        st.rerun()
