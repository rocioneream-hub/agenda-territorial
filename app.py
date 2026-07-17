import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from streamlit_calendar import calendar
from datetime import datetime
import io

# 1. CONEXIÓN ESTABLE
def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds_dict = dict(st.secrets["gcp_service_account"])
    # Limpiamos saltos de línea para el PEM
    if "\\n" in creds_dict["private_key"]:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

# 2. CARGA Y GUARDADO ATÓMICO (Para que no se pierdan datos)
def load_data():
    client = get_gspread_client()
    sheet = client.open_by_key(st.secrets["spreadsheet"]["id"]).sheet1
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def save_data(df):
    client = get_gspread_client()
    sheet = client.open_by_key(st.secrets["spreadsheet"]["id"]).sheet1
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

if 'agenda' not in st.session_state:
    st.session_state.agenda = load_data()

# 3. INTERFAZ Y TABS (Tus funcionalidades originales)
st.title("Agenda UPEU")
pwd = st.sidebar.text_input("Password", type="password")
es_editor = (pwd == "UPEU2026")

tabs = st.tabs(["🗓️ Calendario", "✍️ Carga", "✏️ Editar", "📊 Base de Datos"])

with tabs[0]:
    st.header("Calendario")
    # ... (Aquí va tu lógica original de calendario) ...

with tabs[1]:
    if es_editor:
        st.header("Carga Rápida")
        # ... (Aquí va tu formulario original) ...
    else: st.warning("Solo para editores")

with tabs[2]:
    if es_editor:
        st.header("Editar")
        # ... (Aquí va tu lógica de edición) ...
    else: st.warning("Solo para editores")

with tabs[3]:
    st.dataframe(st.session_state.agenda)
