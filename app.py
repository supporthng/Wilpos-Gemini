import io
import json
import os
import time
from datetime import datetime
import re
import difflib
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
            "ALVAREZ & SANCHEZ": {"nombre": "ALVAREZ & SANCHEZ", "formato_empaque": "formula_tamano_slash_4x6"},
            "GONZALEZ CUESTA": {"nombre": "GONZALEZ CUESTA", "formato_empaque": "caj_pza_estandar"},
            "CND": {"nombre": "CND", "formato_empaque": "universal_extractor"},
            "BEES": {"nombre": "BEES", "formato_empaque": "universal_extractor"}
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
    st.success(f"🟢 **WilPOS Activo** | Catálogo: **{len(master_dict):,}** prods | 🏢 Proveedores: **{len(supps)}** | 🕒 Última act: **{meta.get('ultima_actualizacion', 'Desconocida')}**")

def safe_float(val, default=0.0):
    try: return float(val)
    except (ValueError, TypeError): return default

def safe_int(val, default=1):
    try: return int(val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C (Sin Codigo)"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]: return "S/C (Sin Codigo)"
    return s_val

def get_strict_ean_code_and_name(description, invoice_ean="", supplier_name=""):
    desc_clean = str(description).strip().upper()
    master = st.session_state["master_catalog"]
    memory = st.session_state["barcode_memory"]
    
    if desc_clean in master: return desc_clean, clean_ean_code(master[desc_clean])
    if desc_clean in memory: return desc_clean, clean_ean_code(memory[desc_clean])
        
    master_keys = list(master.keys())
    if master_keys:
        coincidencias = difflib.get_close_matches(desc_clean, master_keys, n=1, cutoff=0.60)
        if coincidencias: return desc_clean, clean_ean_code(master[coincidencias[0]])
            
    s_name = str(supplier_name).upper()
    cleaned_invoice_ean = clean_ean_code(invoice_ean)
    if ("CND" in s_name or "BEES" in s_name) and len(cleaned_invoice_ean) <= 6:
        return desc_clean, "S/C (Sin Codigo)"

    if cleaned_invoice_ean != "S/C (Sin Codigo)": return desc_clean, cleaned_invoice_ean
    return desc_clean, "S/C (Sin Codigo)"

# ==========================================
# REGLA DE EMPAQUE UNIVERSAL REFORZADA (CON EXCEPCIONES FIJAS)
# ==========================================
def parse_empaque_universal(supplier_name="", tamano_txt="", unidad_txt="", descripcion_txt=""):
    combined = f"{str(tamano_txt)} {str(unidad_txt)} {str(descripcion_txt)}".upper()
    
    # Excepciones específicas exactas
    if "CLAMATO" in combined:
        return 24
    if "ALOE PURE" in combined:
        return 12
        
    # 1. Buscar formatos con barra como 16/650, 24/12, 6/473, 12/
    match_slash = re.search(r'\b(24|16|12|6|48|10|20|30)\s*/', combined)
    if match_slash:
        return int(match_slash.group(1))
        
    # 2. Buscar formato matriz como 4X6, 6X4, 4X (LP 4)
    match_nxn = re.search(r'\b(\d+)\s*[xX]\s*(\d+)\b', combined)
    if match_nxn:
        return int(match_nxn.group(1)) * int(match_nxn.group(2))
        
    if re.search(r'\b4\s*[xX]\b', combined) or "LP 4" in combined:
        return 24
        
    # 3. Buscar palabras explícitas de empaque
    match_words = re.search(r'\b(24|16|12|6|48)\s*(BOTS|BOTELLAS|PACK|PZA|UNIDADES|UN)\b', combined)
    if match_words:
        return int(match_words.group(1))
        
    # 4. Reglas específicas para bebidas comunes
    if "GATORADE" in combined and "24" in combined:
        return 24
        
    if "HUACAL" in combined or "PTE. HU" in combined or "BRAHMA LIGHT HU" in combined or "THE ONE HU" in combined:
        if "24" in combined: return 24
        if "16" in combined: return 16
        
    return 1

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Sistema Universal de Empaques</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio("Menú de Navegación", ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📁 Actualizar Catálogo Maestro", "🏢 Perfiles de Proveedores", "📜 Historial de Procesados", "📋 Códigos Almacenados"])

st.sidebar.markdown("---")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

def process_invoice_smart_router(file_obj, file_type, use_paid_gemini=False):
    prompt_detect = "Identifica el nombre comercial del proveedor emisor de esta factura (ej: CND, BEES, ALVAREZ & SANCHEZ, GONZALEZ CUESTA, PRICESMART). Devuelve un JSON puro: {'proveedor': 'NOMBRE'}"
    
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

    supp_mem = st.session_state["supplier_memory"]
    if supplier_detected not in supp_mem and supplier_detected != "GENERAL":
        supp_mem[supplier_detected] = {"nombre": supplier_detected, "formato_empaque": "universal_extractor", "descripcion": "Registrado automaticamente."}
        save_supplier_memory()

    prompt_main = (
        f"Analiza este documento de compra del proveedor '{supplier_detected}' con absoluta precisión. "
        "Para cada renglón extrae rigurosamente en un JSON bajo la clave 'items': "
        "1. 'codigo_ean': código de barras oficial si existe. "
        "2. 'descripcion': nombre exacto del producto con todas sus especificaciones. "
        "3. 'cantidad': cantidad comprada exactamente tal como aparece. "
        "4. 'unidad': unidad de medida exacta impresa en la línea ('UN', 'PC'). "
        "5. 'valor': monto total neto de la línea. "
        "Estructura JSON exacta: "
        '{"proveedor": "' + supplier_detected + '", "items": [{"codigo_ean": "...", "descripcion": "...", "cantidad": 1.0, "unidad": "...", "valor": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
        try:
            file_obj.seek(0)
            response = model.generate_content([image_input, prompt_main])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("
