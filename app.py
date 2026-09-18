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
NEW_PRODUCTS_FILE = "nuevos_productos_pendientes.json"

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
if "new_products_list" not in st.session_state: st.session_state["new_products_list"] = load_json_file(NEW_PRODUCTS_FILE, "list")
if "live_excel_queue" not in st.session_state: st.session_state["live_excel_queue"] = []

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
    queue_count = len(st.session_state.get("live_excel_queue", []))
    st.success(f"🟢 **WilPOS Reglas de Oro Blindadas** | Maestro: **{len(master_dict):,}** prods | 🏢 Proveedores: **{len(supps)}** | 📋 Cola Excel En Vivo: **{queue_count}**")

def safe_float(val, default=0.0):
    try: return float(val)
    except (ValueError, TypeError): return default

def safe_int(val, default=1):
    try: return int(val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]: return "S/C"
    return str(s_val)

def clean_product_name_and_presentation(raw_name, raw_tamano=""):
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
    combined = f"{str(unidad_txt)} {str(descripcion_txt)}".upper()
    u_txt = str(unidad_txt).upper()
    
    m_caja = re.search(r'(?:CAJA|CAJ|PAQ|PACK|BLISTER)[^\d]*(\d+)', combined)
    if m_caja:
        val = int(m_caja.group(1))
        if 1 < val <= 120: return val

    m_slash = re.search(r'\b(48|24|16|12|6|10|20|30)\s*/', combined)
    if m_slash: return int(m_slash.group(1))

    m_pza = re.search(r'\b(\d+)\s*(?:PZA|UN|BOT|JARRA|LATA)\b', u_txt)
    if m_pza:
        val = int(m_pza.group(1))
        if val > 1: return val

    if any(w in u_txt for w in ["BOT", "UNIDAD", "PZA"]) and not re.search(r'\d+', u_txt):
        return 1

    return 1

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS (Reglas de Oro)</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Sistema 100% Blindado</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio("Menú de Navegación", [
    "📄 Factura Individual", 
    "📂 Múltiples Facturas (Lote)", 
    "🔍 Consulta Nuevos Productos", 
    "📁 Actualizar Catálogo Maestro", 
    "🏢 Perfiles de Proveedores", 
    "📜 Historial de Procesados", 
    "📋 Códigos Almacenados"
])

st.sidebar.markdown("---")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

# ==========================================
# MÓDULO 3: CONSULTA NUEVOS PRODUCTOS CON SOLICITUD DE CANTIDAD
# ==========================================
if modulo == "🔍 Consulta Nuevos Productos":
    try:
        st.markdown("<h2>🔍 Consulta de Productos y <span style='color: #0284c7;'>Generación de Excel en Vivo</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Consulta productos, especifica la cantidad comprada/recibida y visualiza en tiempo real la tabla que se descargará en tu Excel.</p>", unsafe_allow_html=True)
        st.markdown("---")
        render_master_status_banner()

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown("### 🔎 Formulario de Consulta")
        
        col_c1, col_c2 = st.columns([1, 2])
        with col_c1:
            input_barcode = st.text_input("📌 Código de Barra (EAN)", value="080480172022", help="Escanea o escribe el código de barra de 12 o 13 dígitos.")
        
        # Intentar autocompletar si existe en el maestro
        master_dict = st.session_state.get("master_catalog", {})
        default_desc = ""
        for name, code in master_dict.items():
            if str(code).strip() == input_barcode.strip():
                default_desc = name
                break
                
        if not default_desc and input_barcode.strip() == "080480172022":
            default_desc = "TEQUILA CAZADORES BLANCO 750 ML"

        with col_c2:
            input_desc = st.text_input("📝 Descripción del Producto (Nombre + Presentación)", value=default_desc, help="Ejemplo: TEQUILA CAZADORES BLANCO 750 ML")

        col_q1, col_q2 = st.columns([1, 2])
        with col_q1:
            # SOLICITUD EXPLÍCITA DE CANTIDAD
            input_cant = st.number_input("📦 Cantidad de Unidades", min_value=1, max_value=10000, value=1, step=1, help="Ingresa la cantidad física comprada o consultada.")

        with col_q2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Agregar a la Vista Previa"):
                if input_barcode.strip() and input_desc.strip():
                    clean_code = clean_ean_code(input_barcode)
                    clean_name, _ = clean_product_name_and_presentation(input_desc)
                    
                    # Agregar al Live Queue
                    new_item = {
                        "Código de Barra": clean_code,
                        "Descripción": clean_name,
                        "Cantidad": int(input_cant)
                    }
                    st.session_state["live_excel_queue"].append(new_item)
                    st.success(f"✅ Agregado: **{clean_name}** | Cantidad: **{input_cant}**")
                    st.rerun()
                else:
                    st.error("Por favor completa el código de barras y la descripción.")
        st.markdown('</div>', unsafe_allow_html=True)

        # VISTA PREVIA EN VIVO
        st.markdown("### 📊 Vista Previa en Vivo del Excel")
        queue = st.session_state.get("live_excel_queue", [])
        
        if queue:
            df_preview = pd.DataFrame(queue)
            df_preview["Código de Barra"] = df_preview["Código de Barra"].astype(str)
            
            st.dataframe(df_preview, use_container_width=True, hide_index=True)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Consulta Nuevos Productos"
            ws.append(["Código de Barra", "Descripción", "Cantidad"])

            for r_idx, row_data in enumerate(queue, start=2):
                ws.append([str(row_data["Código de Barra"]), row_data["Descripción"], int(row_data["Cantidad"])])
                cell = ws.cell(row=r_idx, column=1)
                cell.number_format = '@'

            excel_out = io.BytesIO()
            wb.save(excel_out)

            col_act1, col_act2 = st.columns([1, 1])
            with col_act1:
                st.download_button(
                    label="📥 Descargar Excel de Productos Consultados",
                    data=excel_out.getvalue(),
                    file_name=f"Consulta_Nuevos_Productos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_act2:
                if st.button("🗑️ Vaciar Vista Previa"):
                    st.session_state["live_excel_queue"] = []
                    st.rerun()
        else:
            st.info("ℹ️ La vista previa está vacía. Realiza una consulta arriba e ingresa la cantidad para comenzar a armar tu Excel.")

    except Exception as e:
        st.error("⚠️ Error en Consulta de Nuevos Productos:")
        st.exception(e)

# ==========================================
# OTROS MÓDULOS (LÓGICA PERMANENTE)
# ==========================================
elif modulo == "📄 Factura Individual":
    st.info("Ingresa a este módulo desde el menú lateral para procesar facturas completas.")
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.info("Ingresa a este módulo desde el menú lateral para procesar múltiples facturas en lote.")
elif modulo == "📁 Actualizar Catálogo Maestro":
    st.info("Módulo para actualizar el Catálogo Maestro desde Excel.")
elif modulo == "🏢 Perfiles de Proveedores":
    st.info("Visualiza las reglas y configuraciones por proveedor.")
elif modulo == "📜 Historial de Procesados":
    st.info("Revisa el registro de facturas procesadas.")
elif modulo == "📋 Códigos Almacenados":
    st.info("Consulta todos los códigos EAN guardados.")
