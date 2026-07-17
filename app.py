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

st.set_page_config(layout="wide", page_title="Agenda UPEU", page_icon="🗓️")

# --- CONEXIÓN BLINDADA ---
def obtener_servicio_sheets():
    alcance = ["https://www.googleapis.com/auth/spreadsheets"]
    cred = dict(st.secrets["gcp_service_account"])
    if "\\n" in cred["private_key"]: cred["private_key"] = cred["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(cred, scopes=alcance)
    return build('sheets', 'v4', credentials=creds)

def load_data_from_sheets():
    try:
        servicio = obtener_servicio_sheets()
        res = servicio.spreadsheets().values().get(spreadsheetId=st.secrets["spreadsheet"]["id"], range="A1:Z2000").execute()
        filas = res.get('values', [])
        df = pd.DataFrame(filas[1:], columns=[str(h).strip() for h in filas[0]])
        df['Cantidad de personas estimadas'] = pd.to_numeric(df['Cantidad de personas estimadas'], errors='coerce').fillna(0).astype(int)
        return df
    except: return pd.DataFrame()

def guardar_data(df):
    servicio = obtener_servicio_sheets()
    lista = [df.columns.values.tolist()] + df.astype(str).values.tolist()
    servicio.spreadsheets().values().clear(spreadsheetId=st.secrets["spreadsheet"]["id"], range="A1:Z2000").execute()
    servicio.spreadsheets().values().update(spreadsheetId=st.secrets["spreadsheet"]["id"], range="A1", valueInputOption="USER_ENTERED", body={'values': lista}).execute()

if 'agenda' not in st.session_state: st.session_state.agenda = load_data_from_sheets()

# --- INTERFAZ ---
st.title("Agenda de Planificación Territorial")
password = st.sidebar.text_input("Contraseña de Editor", type="password")
es_editor = (password == "UPEU2026")

tabs = ["🗓️ Calendario", "✍️ Carga", "✏️ Modificar", "📊 Base de Datos"] if es_editor else ["🗓️ Calendario", "📊 Base de Datos"]
selected_tab = st.tabs(tabs)

# Lógica simplificada de tabs (puedes expandir con el contenido previo)
with selected_tab[0]:
    st.header("Planificación")
    events = [{"title": r['Actividad'], "start": r['Fecha']} for _, r in st.session_state.agenda.iterrows() if r['Fecha']]
    calendar(events=events, options={"initialView": "dayGridMonth"})

with selected_tab[-1]:
    st.dataframe(st.session_state.agenda)
    if st.button("🔄 Sincronizar"):
        st.session_state.agenda = load_data_from_sheets()
        st.rerun()

# Si es editor, agregamos la lógica de carga y modificación que tenías antes
if es_editor:
    with selected_tab[1]:
        st.header("Carga Rápida")
        # Aquí va tu formulario de carga anterior...
    with selected_tab[2]:
        st.header("Modificar Actividad")
        # Aquí va tu lógica de edición anterior...
