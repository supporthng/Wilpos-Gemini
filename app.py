import io
import json
import os
import time
import re
import google.generativeai as genai
from PIL import Image
import streamlit as st
import openpyxl
import pandas as pd

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS CSS
# ==========================================
st.set_page_config(
    page_title="WilPOS - Automatizador Inteligente", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp {
        background-color: #f8fafc;
        color: #1e293b;
        font-family: 'Inter', sans-serif;
    }
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    h1, h2, h3 {
        color: #0f172a;
        font-weight: 700;
    }
    .card-container {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    .stButton>button {
        background: #0284c7;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background: #0369a1;
        color: white;
    }
    .stDownloadButton>button {
        background: #10b981;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.5rem 1.2rem;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .stDownloadButton>button:hover {
        background: #059669;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------
# BÚSQUEDA UNIVERSAL DE CLAVES API
# ------------------------------------------
api_key_candidates = [
    st.secrets.get("GEMINI_API_KEY"),
    st.secrets.get("GOOGLE_API_KEY"),
    st.secrets.get("GEMINI_API_KEY_PAID"),
    st.secrets.get("GEMINI_API_KEY_1"),
    os.environ.get("GEMINI_API_KEY"),
    os.environ.get("GOOGLE_API_KEY"),
    os.environ.get("GEMINI_API_KEY_PAID"),
    os.environ.get("GEMINI_API_KEY_1")
]
ACTIVE_GEMINI_KEY = next((k for k in api_key_candidates if k and str(k).strip()), None)

# Diccionario maestro blindado para Álvarez & Sánchez
MASTER_ALVAREZ_SANCHEZ = {
    "VINO TINTO RESERVA CUNE (D.O.RIOJA)": "8410591003045",
    "VINO TINTO MERLOT VIÑA TARAPACA": "7804340909534",
    "VINO TINTO RESERVA CAB SAUV TARAPACA": "7804340909039",
    "VINO TINTO RESERVA CARMENERE TARAPACA": "7804340909010",
    "VINO TINTO RESERVA MERLOT TARAPACA": "78043409041635",
    "VINO TINTO RED BLEND JUAN GIL(JUMILLA)": "8437010482341",
    "VINO TTO ET. AMARILLA JUAN GIL (JUMILLA)": "8437010482297",
    "VINO TTO AZUL JUAN GIL (JUMILLA)": "8437010482273",
    "VINO TTO PLATA JUAN GIL (JUMILLA)": "8437010482280",
    "VINO TINTO MERLOT CALIFORNIA JOSH": "85000020709",
    "VINO TTO CAB SAUV BOURBON RESERV JOSH": "85000020747",
    "VINO TTO CAB SAUV RESERV JOSH": "85000020754",
    "VINO TINTO PINOT NOIR 689 CELLARS": "85000001942",
    "VINO TINTO SIX EIGHT NINE": "051497322618",
    "WHISKY ESCOSÉS MALTA 12 AÑOS GLEN GRANT": "8000040630269",
    "VODKA INFUSIONS CITRUS SKYY": "72105927504",
    "VODKA INFUSIONS RASPBERRY SKYY": "72105920203",
    "VODKA SKYY": "721059007504"
}

def safe_float(val, default=0.0):
    try:
        return float(val)
    except (ValueError, TypeError):
        return default

def safe_int(val, default=1):
    try:
        return int(val)
    except (ValueError, TypeError):
        return default

def round_to_nearest_5(x):
    return float(round(round(x / 5) * 5))

def clean_barcode(code_val):
    if not code_val:
        return "S/C (Sin Código)"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'):
        s_val = s_val[:-2]
    digits = re.sub(r'\D', '', s_val)
    if not digits:
        return "S/C (Sin Código)"
    return digits

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Automatizador Inteligente (Flash Free)</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

def parse_empaque_exact(desc, unidad_txt):
    d = str(desc).upper()
    u = str(unidad_txt).upper()
    if "BOT" in u and "VODKA" in d:
        return 1
    if "GLEN GRANT" in d or "WHISKY" in d:
        return 12
    if "CAJA" in u:
        return 12 if "JOSH" in d or "CUNE" in d or "TARAPACA" in d else 6
    return 12

def process_invoice_gemini_flash(file_obj, file_type):
    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini (Configura GEMINI_API_KEY en st.secrets)"

    prompt_text = (
        "Analiza esta factura con precisión milimétrica. "
        "Extrae cada renglón de la tabla con: "
        "1. 'descripcion': texto exacto de la columna 'DESCRIPCION'. "
        "2. 'cantidad': número de la columna 'CANTIDAD'. "
        "3. 'unidad': 'CAJA' o 'BOT.'. "
        "4. 'tamano': texto exacto de la columna 'TAMAÑO'. "
        "5. 'precio_lista': número exacto de la columna 'PRECIO'. "
        "6. 'descuento_porcentaje': porcentaje exacto de la columna 'COM.'. "
        "Devuelve un JSON puro bajo la clave 'items': "
        '{"items": [{"descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            # Modelo corregido y actualizado a gemini-2.5-flash
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            
            response = model.generate_content([image_input, prompt_text])
            
            if not response or not response.text:
                raise ValueError("La API de Gemini devolvió una respuesta vacía.")
                
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("
