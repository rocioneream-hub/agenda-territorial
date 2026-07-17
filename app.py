import os
import json
from google.oauth2.service_account import Credentials

def get_gspread_client():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # Intentamos leer el JSON completo desde una variable de entorno llamada 'GCP_JSON'
    # Esto es mucho más robusto que los Secrets de Streamlit
    json_str = os.environ.get('GCP_JSON')
    if not json_str:
        # Fallback: si no está en entorno, lo busca en los Secrets
        json_str = json.dumps(dict(st.secrets["gcp_service_account"]))
    
    creds_dict = json.loads(json_str)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)
