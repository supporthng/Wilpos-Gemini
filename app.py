import io
import json
import os
import time
import hashlib
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

openai_key_candidates = [
    st.secrets.get("OPENAI_API_KEY"),
    os.environ.get("OPENAI_API_KEY")
]
ACTIVE_OPENAI_KEY = next((k for k in openai_key_candidates if k and str(k).strip()), None)

BARCODE_MEMORY_FILE = "codigos_escaneados_memoria.json"

def load_json_file(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_json_file(filepath, data_dict):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando {filepath}: {e}")

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    t = text.upper()
    t = t.replace('SIX EIGHT NINE', '689').replace('SIX-EIGHT-NINE', '689')
    t = t.replace('COGÑA', 'COGNAC').replace('COGÑAC', 'COGNAC').replace('CONGNAC', 'COGNAC')
    t = t.replace('VSOP', 'V.S.O.P').replace('V S O P', 'V.S.O.P')
    t = t.replace('VS ', 'VERY SPECIAL ').replace(' VS', ' VERY SPECIAL')
    t = t.replace('GIN ', 'GINEBRA ').replace(' GIN', ' GINEBRA')
    t = t.replace(' 5CL', ' 50 ML').replace(' 5 CL', ' 50 ML').replace('5CL', '50 ML')
    t = t.replace(' 75CL', ' 750 ML').replace(' 75 CL', ' 750 ML').replace('75CL', '750 ML')
    t = t.replace(' 70CL', ' 700 ML').replace(' 70 CL', ' 700 ML').replace('70CL', '700 ML')
    t = t.replace('PTE.', 'PRESIDENTE').replace('HU', '').replace('CJ', '').replace('BOT.', '')
    for ch in ['/', '-', ',', '.', '(', ')', '%', '+', '"', "'", 'º']:
        t = t.replace(ch, ' ')
    return " ".join(t.split())

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

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Automatizador Inteligente</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📋 Ver Códigos Almacenados"]
)

# ==========================================
# MOTOR MAESTRO INTELIGENTE
# ==========================================
def match_official_barcode(item_description):
    raw_name = str(item_description).strip().upper()
    b_mem = st.session_state["barcode_memory"]
    
    if not b_mem:
        return "S/C (Sin Código)", raw_name, "⚠️ Memoria Vacía"

    if raw_name in b_mem:
        return clean_barcode(b_mem[raw_name]), raw_name, "Maestro Exacto"

    norm_input = normalize_text(raw_name)
    input_tokens = set(norm_input.split())

    best_score = 0.0
    best_code = "S/C (Sin Código)"
    best_name = raw_name

    for master_name, code in b_mem.items():
        norm_master = normalize_text(master_name)
        master_tokens = set(norm_master.split())
        
        if not master_tokens:
            continue

        common_tokens = input_tokens.intersection(master_tokens)
        token_score = len(common_tokens) / max(len(input_tokens), len(master_tokens))
        seq_ratio = difflib.SequenceMatcher(None, norm_input, norm_master).ratio()
        combined_score = (token_score * 0.65) + (seq_ratio * 0.35)
        
        if combined_score > best_score:
            best_score = combined_score
            best_code = code
            best_name = master_name

    if best_score >= 0.28:
        return clean_barcode(best_code), best_name, f"Smart Match ({best_score:.2f})"

    return "S/C (Sin Código)", raw_name, "⚠️ Sin Coincidencia"

def parse_empaque_from_description(item_desc, ai_empaque):
    emp_ai = safe_int(ai_empaque, 1)
    if emp_ai > 1:
        return emp_ai
        
    d = str(item_desc).upper()
    match_slash = re.search(r'(\d+)\s*(?:/|X|x)\s*([\d\.]+)', d)
    if match_slash:
        val1 = int(match_slash.group(1))
        val2 = float(match_slash.group(2))
        if val1 in [6, 12, 24, 20, 30, 48]:
            return val1
        elif val2 in [6, 12, 24, 20, 30]:
            return int(val2)
        return val1

    if any(b in d for b in ["PRESIDENTE", "MICHELOB", "COORS", "BRAHMA", "CORONA", "STELLA", "HEINEKEN", "BECKS", "ESTRELLA", "PERONI"]):
        if "22OZ" in d or "650ML" in d or "GRANDE" in d:
            return 12
        return 24
        
    if any(w in d for w in ["VINO", "WHISKY", "VODKA", "TEQUILA", "RON", "RUM", "GIN", "COGNAC", "LICOR", "FIREBALL", "CAMPARI", "AMARETTO", "KAHLUA", "MIDORI", "STOLICHNAYA", "JOSH", "JUAN GIL", "TARAPACA", "GLENLIVET", "BUCHANAN", "OLD PARR"]):
        if "6" in d or "6/" in d:
            return 6
        return 12

    if any(bev in d for bev in ["GATORADE", "ALOE", "CLAMATO", "REDBULL", "MONSTER", "COCA", "PEPSI", "AGUA", "OCEANSPRAY", "FOURLOKO", "THEONE", "SCHWEPPES", "FEVER TREE"]):
        return 12

    return 1

def process_invoice_with_ai(file_obj, file_type, use_openai_fallback=False):
    if use_openai_fallback:
        if not ACTIVE_OPENAI_KEY:
            return None, "Falta clave API de OpenAI"
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"

        prompt_text = (
            "Analiza esta factura COMPLETAMENTE de arriba a abajo. Extrae TODOS los ítems de la tabla sin omitir ninguno. "
            "Para cada ítem, extrae estrictamente: 'descripcion', 'cantidad', 'empaque', "
            "'precio_lista' (el precio unitario o de caja indicado en la columna PRECIO antes de descuento), "
            "y 'descuento_porcentaje' (el porcentaje de descuento indicado en COM. o DESC., ejemplo: 10 para 10%, o 0). "
            "Devuelve un JSON puro con esta estructura exacta: "
            '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "precio_lista": 0.0, "descuento_porcentaje": 0.0}]}. '
            "Respuesta JSON pura sin texto adicional ni markdown."
        )
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI)"
        except Exception as e:
            return None, str(e)

    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini"

    prompt_text = (
        "Analiza esta factura COMPLETAMENTE de arriba a abajo. Extrae TODOS los ítems de la tabla sin omitir ninguno. "
        "Para cada ítem, extrae estrictamente: 'descripcion', 'cantidad', 'empaque', "
        "'precio_lista' (el precio unitario o de caja indicado en la columna PRECIO antes de descuento), "
        "y 'descuento_porcentaje' (el porcentaje de descuento indicado en COM. o DESC., ejemplo: 10 para 10%, o 0). "
        "Devuelve un JSON puro con esta estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "precio_lista": 0.0, "descuento_porcentaje": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("
