import streamlit as st
import pandas as pd
from streamlit_calendar import calendar
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, time, date
import re
import os
import io

# Importación segura de docx
try:
    import docx
    from docx.shared import Pt, Inches
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    LIBRERIA_DOCX_LISTA = True
except ImportError:
    LIBRERIA_DOCX_LISTA = False

# ==========================================
# 1. CONFIGURACIÓN E IDENTIDAD VISUAL
# ==========================================
st.set_page_config(layout="wide", page_title="Calendario territorial UPEU", page_icon="🗓️")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Figtree:ital,wght@0,300..900;1,300..900&display=swap');
    html, body, [class*="css"], .stMarkdown, p, div { font-family: 'Figtree', sans-serif !important; }
    .stApp, .main { background-color: #E8E8E8 !important; }
    div[data-testid="stForm"] { background-color: #FFFFFF !important; border-radius: 8px !important; padding: 30px !important; border: 1px solid #D1D5DB !important; }
    h1 { color: #000000 !important; font-weight: 800 !important; }
    h2, h3 { color: #007BE0 !important; }
    .hashtag-gestion { color: #6AC64F !important; font-weight: 800; font-size: 1.1rem; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CONEXIÓN BLINDADA (La versión que no falla)
# ==========================================
def obtener_servicio_sheets():
    alcance = ["https://www.googleapis.com/auth/spreadsheets"]
    credenciales_dict = dict(st.secrets["gcp_service_account"])
    # Limpieza necesaria para el PEM
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
        encabezados = [str(h).strip() for h in filas[0]]
        df = pd.DataFrame(filas[1:], columns=encabezados)
        return df
    except Exception as e:
        st.error(f"Error cargando base: {e}")
        return pd.DataFrame()

# [AQUÍ MANTENEMOS TODAS TUS FUNCIONES: guardar_todo_en_sheets, limpiar_fecha_para_calendario, crear_reporte_word, etc...]
# (Nota: He resumido por espacio, pero mantenés toda la lógica que tenías antes en los tabs)

# ==========================================
# 3. INTERFAZ (Tabs, Formulario de carga, Calendar, Reportes)
# ==========================================
# ... [Aquí va todo tu código de tabs, calendario y botones que ya tenías] ...

# Recordá: La lógica de Tabs y Botones no cambió, solo la función de conexión que es la que te causaba el conflicto.
