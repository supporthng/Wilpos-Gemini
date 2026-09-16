import io
import json
import os
import time
import re
import difflib
import google.generativeai as genai
from openai import OpenAI
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
gemini_paid_candidates = [
    st.secrets.get("GEMINI_API_KEY_PAID"),
    os.environ.get("GEMINI_API_KEY_PAID")
]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [
    st.secrets.get("GEMINI_API_KEY"),
    st.secrets.get("GOOGLE_API_KEY"),
    st.secrets.get("GEMINI_API_KEY_1"),
    os.environ.get("GEMINI_API_KEY"),
    os.environ.get("GOOGLE_API_KEY"),
    os.environ.get("GEMINI_API_KEY_1")
]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

openai_key_candidates = [
    st.secrets.get("OPENAI_API_KEY"),
    os.environ.get("OPENAI_API_KEY")
]
ACTIVE_OPENAI_KEY = next((k for k in openai_key_candidates if k and str(k).strip()), None)

BARCODE_MEMORY_FILE = "codigos_escaneados_memoria.json"

def load_json_file(filepath):
    data = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    
    # Códigos oficiales fijos verificados en factura
    data["WHISKY ESCOCES MALTA 12 AÑOS GLEN GRANT"] = "8000040630269"
    data["VINO TINTO SIX EIGHT NINE 689"] = "051497322618"
    data["VODKA SKYY"] = "721059007504"
    data["VODKA INFUSIONS CITRUS SKYY"] = "721059627504"
    data["VODKA INFUSIONS RASPBERRY SKYY"] = "721059637503"
    return data

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = {}

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
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]:
        return "S/C (Sin Código)"
    return s_val

def get_resolved_barcode(extracted_code, description):
    cleaned = clean_barcode(extracted_code)
    desc_upper = str(description).strip().upper()
    
    if cleaned != "S/C (Sin Código)":
        st.session_state["barcode_memory"][desc_upper] = cleaned
        return cleaned
    
    if desc_upper in st.session_state["master_catalog"]:
        return st.session_state["master_catalog"][desc_upper]
        
    if desc_upper in st.session_state["barcode_memory"]:
        return st.session_state["barcode_memory"][desc_upper]
        
    return "S/C (Sin Código)"

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Automatizador Inteligente</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📁 Cargar Archivo Maestro", "📋 Ver Códigos Almacenados"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Configuración de API")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

def parse_empaque_from_tamano(tamano_txt, unidad_txt):
    u = str(unidad_txt).strip().upper()
    if "BOT" in u:
        return 1
    t = str(tamano_txt).strip()
    match_t = re.search(r'^(\d+)\s*/', t)
    if match_t:
        return int(match_t.group(1))
    return 12

def process_invoice_exact_18(file_obj, file_type, use_paid_gemini=False, use_openai_fallback=False):
    prompt_text = (
        "Analiza esta factura de Álvarez & Sánchez con extrema precisión horizontal. La tabla tiene filas numeradas. "
        "Debes asegurar que el 'codigo_barras' de la columna izquierda esté estrictamente alineado con la 'descripcion' de esa misma línea horizontal exacta, sin desplazar los códigos hacia arriba ni hacia abajo. "
        "Para cada renglón extrae exactamente: "
        "1. 'codigo_barras': el número exacto de la columna 'CODIGO DE BARRAS' que se encuentra en la misma fila horizontal de la descripción. "
        "2. 'descripcion': el texto exacto de la columna 'DESCRIPCION'. "
        "3. 'cantidad': número de la columna 'CANTDAD'. "
        "4. 'unidad': 'CAJA' o 'BOT.'. "
        "5. 'tamano': texto exacto de la columna 'TAMAÑO' (ej: 12/75 CL., 6/70 CL., 75 CL.). "
        "6. 'precio_lista': número exacto de la columna 'PRECIO'. "
        "7. 'descuento_porcentaje': porcentaje de la columna 'COM.' (ej: 10). "
        "Devuelve un JSON puro con un arreglo exacto de objetos bajo la clave 'items': "
        '{"items": [{"codigo_barras": "...", "descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini:
        active_key = ACTIVE_GEMINI_FREE_KEY

    if not active_key and use_openai_fallback:
        use_openai_fallback = True

    if use_openai_fallback and not active_key:
        if not ACTIVE_OPENAI_KEY:
            return None, "Faltan claves API válidas (Gemini / OpenAI)"
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI gpt-4o)"
        except Exception as e:
            return None, str(e)

    for intento in range(2):
        try:
            genai.configure(api_key=active_key)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower():
                if not use_paid_gemini and ACTIVE_GEMINI_PAID_KEY:
                    active_key = ACTIVE_GEMINI_PAID_KEY
                    continue
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
            
    if use_openai_fallback and ACTIVE_OPENAI_KEY:
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("
