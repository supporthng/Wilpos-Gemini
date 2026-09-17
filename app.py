import io
import json
import os
import time
from datetime import datetime
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
    page_title="WilPOS - Sistema de Inventario Inteligente", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e2e8f0; }
    h1, h2, h3 { color: #0f172a; font-weight: 700; }
    .card-container { background-color: #ffffff; border: 1px solid #e2e8f0; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); margin-bottom: 20px; }
    [data-testid="stMetricValue"] { font-size: 1.25rem !important; font-weight: 700; white-space: nowrap; overflow: visible; }
    .stButton>button { background: #0284c7; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.5rem; font-weight: 600; font-size: 1rem; transition: all 0.2s ease; }
    .stButton>button:hover { background: #0369a1; color: white; }
    .stDownloadButton>button { background: #10b981; color: white; border: none; border-radius: 8px; padding: 0.5rem 1.2rem; font-weight: 600; transition: all 0.2s ease; }
    .stDownloadButton>button:hover { background: #059669; color: white; }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------
# BÚSQUEDA UNIVERSAL DE CLAVES API
# ------------------------------------------
gemini_paid_candidates = [st.secrets.get("GEMINI_API_KEY_PAID") if "GEMINI_API_KEY_PAID" in st.secrets else None, os.environ.get("GEMINI_API_KEY_PAID")]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else None, st.secrets.get("GOOGLE_API_KEY") if "GOOGLE_API_KEY" in st.secrets else None, os.environ.get("GEMINI_API_KEY"), os.environ.get("GOOGLE_API_KEY")]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

BARCODE_MEMORY_FILE = "codigos_escaneados_memoria.json"
MASTER_CATALOG_FILE = "catalogo_maestro_sistema.json"
MASTER_META_FILE = "catalogo_maestro_meta.json"
SUPPLIER_MEMORY_FILE = "proveedores_formatos_memoria.json"
HISTORY_FILE = "historial_procesados.json"

def load_json_file(filepath, default_type="dict"):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if default_type == "list" and isinstance(data, list): return data
                if default_type == "dict" and isinstance(data, dict): return data
        except Exception:
            pass
    return [] if default_type == "list" else {}

def save_json_file(filepath, data):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

if "barcode_memory" not in st.session_state: st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE, "dict")
if "master_catalog" not in st.session_state: st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE, "dict")
if "master_meta" not in st.session_state: st.session_state["master_meta"] = load_json_file(MASTER_META_FILE, "dict")
if "processing_history" not in st.session_state: st.session_state["processing_history"] = load_json_file(HISTORY_FILE, "list")

if "supplier_memory" not in st.session_state:
    loaded_suppliers = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    if not loaded_suppliers:
        loaded_suppliers = {
            "ALVAREZ & SANCHEZ": {"nombre": "ALVAREZ & SANCHEZ", "formato_empaque": "regla_oro_blindada"},
            "GONZALEZ CUESTA": {"nombre": "GONZALEZ CUESTA", "formato_empaque": "regla_oro_blindada"},
            "CENTRO DE DISTRIBUCION CHRISTIAN": {"nombre": "CENTRO DE DISTRIBUCION CHRISTIAN", "formato_empaque": "regla_oro_blindada"}
        }
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_suppliers)
    st.session_state["supplier_memory"] = loaded_suppliers

def save_master_to_file(): save_json_file(MASTER_CATALOG_FILE, st.session_state["master_catalog"])
def save_meta_to_file(timestamp_str, count):
    meta = {"ultima_actualizacion": timestamp_str, "total_productos": count}
    st.session_state["master_meta"] = meta
    save_json_file(MASTER_META_FILE, meta)
def save_supplier_memory(): save_json_file(SUPPLIER_MEMORY_FILE, st.session_state["supplier_memory"])

def add_to_history(entry):
    hist = st.session_state.get("processing_history", [])
    if not isinstance(hist, list): hist = []
    hist.insert(0, entry)
    st.session_state["processing_history"] = hist
    save_json_file(HISTORY_FILE, hist)

def render_master_status_banner():
    master_dict = st.session_state["master_catalog"]
    meta = st.session_state.get("master_meta", {})
    supps = st.session_state.get("supplier_memory", {})
    st.success(f"🟢 **WilPOS Reglas de Oro Blindadas** | Maestro: **{len(master_dict):,}** prods | 🏢 Proveedores: **{len(supps)}**")

def safe_float(val, default=0.0):
    try: return float(val)
    except (ValueError, TypeError): return default

def safe_int(val, default=1):
    try: return int(val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): return float(round(round(x / 5) * 5))

# ==========================================
# REGLAS DE ORO: LIMPIEZA Y CÁLCULO BLINDADO
# ==========================================
def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]: return "S/C"
    return str(s_val)

def clean_product_name_and_presentation(raw_name, raw_tamano=""):
    """Regla 5 de Oro: Nombres limpios combinando Nombre + Presentación (ej. CHIVAS REGAL 25YO 700 ML)."""
    name = str(raw_name).strip()
    name = re.sub(r'^\d+[\s-]*', '', name)
    name = re.sub(r'^\[.*?\]\s*', '', name)
    
    presentation = str(raw_tamano).strip().upper()
    if not presentation or presentation in ["NAN", "NONE", ""]:
        match_pres = re.search(r'\b(\d+\s*/\s*[\d\.]+\s*(?:CL|ML|L|LT|OZ)|\d+\s*(?:CL|ML|L|LT|OZ))\b', name, re.IGNORECASE)
        if match_pres:
            presentation = match_pres.group(1).upper()
            name = name.replace(match_pres.group(1), '')
            
    name = re.sub(r'\s+', ' ', name).strip().upper()
    presentation = re.sub(r'\s+', ' ', presentation).strip().upper()
    
    if presentation and presentation not in name:
        clean_full_name = f"{name} {presentation}"
    else:
        clean_full_name = name
        
    return clean_full_name, presentation

def get_flexible_master_barcode(clean_name, clean_pres=""):
    """Reglas 2, 3 y 6 de Oro: Cruce inteligente tolerante a variaciones, abreviaciones y faltas ortográficas."""
    master_dict = st.session_state.get("master_catalog", {})
    if not master_dict: return "S/C"
    
    full_query = f"{clean_name} {clean_pres}".strip()
    if full_query in master_dict:
        return clean_ean_code(master_dict[full_query])
    if clean_name in master_dict:
        return clean_ean_code(master_dict[clean_name])
        
    query_words = [w for w in re.findall(r'\w+', full_query.upper()) if len(w) > 2]
    if not query_words: return "S/C"
    
    best_code = "S/C"
    max_matches = 0
    
    for m_name, m_code in master_dict.items():
        m_upper = str(m_name).upper()
        master_words = [w for w in re.findall(r'\w+', m_upper) if len(w) > 2]
        
        matches = 0
        for qw in query_words:
            for mw in master_words:
                if qw == mw or (len(qw) >= 4 and (qw in mw or mw in qw)):
                    matches += 1
                    break
                    
        if matches >= 2 and matches > max_matches:
            max_matches = matches
            best_code = clean_ean_code(m_code)
            
    return best_code

def parse_empaque_blindado(unidad_txt="", descripcion_txt="", tamano_txt=""):
    """Regla 4 de Oro (Reforzada): Detección inequívoca de empaques separada del tamaño."""
    combined = f"{str(unidad_txt)} {str(descripcion_txt)}".upper()
    u_txt = str(unidad_txt).upper()
    
    # 1. Patrón explícito de caja/paquete con número (ej: CAJA-24, CAJ 6, PAQ 12)
    m_caja = re.search(r'(?:CAJA|CAJ|PAQ|PACK|BLISTER)[^\d]*(\d+)', combined)
    if m_caja:
        val = int(m_caja.group(1))
        if 1 < val <= 120: return val

    # 2. Patrón de división tipo 6/75 CL o 12/750 ML en la unidad o descripción
    m_slash = re.search(r'\b(48|24|16|12|6|10|20|30)\s*/', combined)
    if m_slash: return int(m_slash.group(1))

    # 3. Patrón de piezas por unidad (ej: "6 PZA", "12 UN")
    m_pza = re.search(r'\b(\d+)\s*(?:PZA|UN|BOT|JARRA|LATA)\b', u_txt)
    if m_pza:
        val = int(m_pza.group(1))
        if val > 1: return val

    # 4. Si dice explícitamente BOT o UNIDAD suelta sin número de caja
    if any(w in u_txt for w in ["BOT", "UNIDAD", "PZA"]) and not re.search(r'\d+', u_txt):
        return 1

    return 1

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS (Reglas de Oro)</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Sistema 100% Blindado</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio("Menú de Navegación", ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📁 Actualizar Catálogo Maestro", "🏢 Perfiles de Proveedores", "📜 Historial de Procesados", "📋 Códigos Almacenados"])

st.sidebar.markdown("---")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

def process_invoice_smart_router(file_obj, file_type, use_paid_gemini=False):
    prompt_detect = "Identifica el nombre comercial del proveedor emisor de esta factura (ej: CND, BEES, ALVAREZ & SANCHEZ, GONZALEZ CUESTA, CENTRO DE DISTRIBUCION CHRISTIAN). Devuelve un JSON puro: {'proveedor': 'NOMBRE'}"
    
    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini: active_key = ACTIVE_GEMINI_FREE_KEY

    genai.configure(api_key=active_key if active_key else ACTIVE_GEMINI_FREE_KEY)
    model = genai.GenerativeModel('gemini-3.6-flash')
    
    file_obj.seek(0)
    file_bytes = file_obj.read()
    image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))

    supplier_detected = "GENERAL"
    try:
        resp_det = model.generate_content([image_input, prompt_detect])
        txt_det = resp_det.text.strip()
        if txt_det.startswith("```json"): txt_det = txt_det[7:]
        if txt_det.endswith("```"): txt_det = txt_det[:-3]
        supplier_detected = str(json.loads(txt_det.strip()).get("proveedor") or "GENERAL").upper().strip()
    except Exception:
        supplier_detected = "GENERAL"

    if supplier_detected and supplier_detected != "GENERAL":
        supps = st.session_state["supplier_memory"]
        if supplier_detected not in supps:
            supps[supplier_detected] = {"nombre": supplier_detected, "formato_empaque": "regla_oro_blindada"}
            st.session_state["supplier_memory"] = supps
            save_supplier_memory()

    prompt_main = (
        f"Analiza este documento de compra del proveedor '{supplier_detected}' bajo las Reglas de Oro. "
        "Extrae ABSOLUTAMENTE TODOS LOS RENGLONES/PRODUCTOS que aparecen en la factura, sin duplicar artificialmente ninguna línea y sin omitir ninguna. "
        "Para cada renglón extrae rigurosamente en un JSON bajo la clave 'items': "
        "1. 'descripcion': nombre del producto. "
        "2. 'tamano': presentación o tamaño de la botella individual (ej: '750 ML', '70 CL', '1 LT'). "
        "3. 'cantidad': cantidad comprada de cajas o unidades (ej: 1.0). "
        "4. 'unidad': unidad de medida o empaque impreso (ej: '6 PZA', 'Caja-24', 'BOT'). "
        "5. 'valor_con_itbis': monto TOTAL INCLUYENDO ITBIS que aparece en la línea (importe final con impuestos y descuentos aplicados). "
        "6. 'descuento_monto': monto del descuento aplicado a esta línea (si existe, ej: 0.0). "
        "Estructura JSON exacta: "
        '{"proveedor": "' + supplier_detected + '", "items": [{"descripcion": "...", "tamano": "...", "cantidad": 1.0, "unidad": "...", "valor_con_itbis": 0.0, "descuento_monto": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
        try:
            file_obj.seek(0)
            response = model.generate_content([image_input, prompt_main])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("
