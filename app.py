import io
import json
import os
import time
import unicodedata
from datetime import datetime
import re
import google.generativeai as genai
from PIL import Image
import streamlit as st
import openpyxl
import pandas as pd

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="WilPOS - Sistema Multi-Proveedor Maestro", 
    page_icon="⚡", 
    layout="wide"
)

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    .card-container { background-color: #ffffff; border: 1px solid #e2e8f0; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); margin-bottom: 20px; }
    .stButton>button { background: #0284c7; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.5rem; font-weight: 600; }
    .stButton>button:hover { background: #0369a1; color: white; }
    .stDownloadButton>button { background: #10b981; color: white; border: none; border-radius: 8px; padding: 0.5rem 1.2rem; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

gemini_key = (
    st.secrets.get("GEMINI_API_KEY_PAID") or 
    st.secrets.get("GEMINI_API_KEY") or 
    os.environ.get("GEMINI_API_KEY_PAID") or 
    os.environ.get("GEMINI_API_KEY")
)

if gemini_key:
    genai.configure(api_key=gemini_key)

SUPPLIER_MEMORY_FILE = "proveedores_formatos_memoria.json"
MASTER_CATALOG_FILE = "catalogo_maestro_sistema.json"

def load_json_file(filepath, default_type="dict"):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if default_type == "dict":
                    if isinstance(data, dict):
                        return {str(k).upper().strip(): v for k, v in data.items() if k}
                    return {}
                elif default_type == "list":
                    if isinstance(data, list):
                        return data
                    return []
        except Exception:
            pass
    return {} if default_type == "dict" else []

def save_json_file(filepath, data):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

# ==========================================
# GESTIÓN DE PERFILES Y CATÁLOGO
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "BEPENSA DOMINICANA SA": {
            "nombre": "BEPENSA DOMINICANA SA",
            "tipo_formato": "factura_tique_bepensa",
            "instruccion_prompt": "Analiza esta página de la factura de BEPENSA DOMINICANA SA renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario' (precio neto del paquete sin ITBIS), 'descuento_porcentaje' y 'monto_neto' (subtotal de la línea sin ITBIS)."
        },
        "EL CATADOR": {
            "nombre": "EL CATADOR",
            "tipo_formato": "factura_cajas_descuento",
            "instruccion_prompt": "Analiza esta página de la factura de EL CATADOR renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "PRICESMART": {
            "nombre": "PRICESMART",
            "tipo_formato": "factura_tique_unidades",
            "instruccion_prompt": "Analiza esta página del comprobante de PRICESMART renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario' y 'monto_neto'."
        },
        "CENTRO DE DISTRIBUCION CRISTIAN": {
            "nombre": "CENTRO DE DISTRIBUCION CRISTIAN",
            "tipo_formato": "pos_cajas_unidades",
            "instruccion_prompt": "Analiza esta página de CENTRO DE DISTRIBUCION CRISTIAN renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_desglose_descuentos",
            "instruccion_prompt": "Analiza esta página de la factura de ALVAREZ & SANCHEZ renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        }
    }
    if not loaded_supps or "BEPENSA DOMINICANA SA" not in loaded_supps:
        loaded_supps.update(default_profiles)
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    loaded_master = load_json_file(MASTER_CATALOG_FILE, "dict")
    if not loaded_master:
        base_defaults = {
            "ANTIOQUEÑO TAPA ROJA 750 ML": "7702131234567",
            "OLD PARR 12 AÑOS 750ML": "7804300120986",
            "FRONTERA SAUVIGNON BLANC C Y T 750ML": "051497455286"
        }
        loaded_master = base_defaults
        save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

if "paginas_procesadas_historial" not in st.session_state:
    st.session_state["paginas_procesadas_historial"] = set()

if "duplicados_confirmados_sesion" not in st.session_state:
    st.session_state["duplicados_confirmados_sesion"] = set()

def safe_float(val, default=0.0):
    try:
        if val is None: return default
        s_val = str(val).replace('$', '').replace('%', '').replace(',', '').strip()
        return float(s_val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): 
    return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    s_val = re.sub(r'\D', '', s_val)
    if 7 <= len(s_val) <= 14: 
        return str(s_val)
    return "S/C"

def normalizar_texto(texto):
    if not texto: return ""
    t = str(texto).upper().strip()
    t = ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
    t = re.sub(r'([A-Z])(\d)', r'\1 \2', t)
    t = re.sub(r'(\d)([A-Z])', r'\1 \2', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def limpiar_nombre_y_extraer_presentacion(proveedor_activo, descripcion_raw, tamano_raw=""):
    t_norm = normalizar_texto(descripcion_raw)
    t_tam = normalizar_texto(tamano_raw)
    prov_up = normalizar_texto(proveedor_activo)
    
    combined_raw = f"{t_norm} {t_tam}"
    
    m_med = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|CL))', combined_raw)
    presentacion = m_med.group(1).replace(" ", "") if m_med else ("750ML" if "EL CATADOR" in prov_up else "UN")
    if "75CL" in presentacion:
        presentacion = "750ML"

    desc_limpia = re.sub(r'\b(CAJA\s*\d*|CJ\s*\d*\s*BOT|\d+\s*X\s*\d+\s*(?:ML|CL|L)|\d+/\s*\d+\s*(?:ML|CL|L|750ML)|750\s*ML|75\s*CL|750ML|75CL|\d+OZ|\d+\s*PZAS|\d+\s*PZA|\d+P\b|\b4P\b|\b12P\b)\b', '', t_norm)
    desc_limpia = re.sub(r'\s+', ' ', desc_limpia).strip()
    
    if presentacion != "UN" and presentacion not in desc_limpia:
        desc_limpia = f"{desc_limpia} {presentacion}".strip()
    
    return desc_limpia, presentacion

def parse_empaque_proveedor(proveedor_nombre, unidad_txt="", tamano_txt="", descripcion_txt=""):
    combined = normalizar_texto(f"{unidad_txt} {tamano_txt} {descripcion_txt}")
    
    m_pzas = re.search(r'(?:(\d+)\s*(?:PZAS|PZA|BOT|UNIDADES|UN|CAJA|CJ|BOX))|(?:(?:PZAS|PZA|BOT|CAJA|CJ|BOX)[\s\-]*(\d+))', combined)
    if m_pzas:
        val = int(m_pzas.group(1) or m_pzas.group(2))
        if val > 0: return val

    m_pack = re.search(r'(?:LT\s*)?(\d+)\s*P\b', combined)
    if m_pack:
        val = int(m_pack.group(1))
        if val > 0: return val

    for num in [12, 24, 6, 4]:
        if f" {num} " in combined or combined.endswith(f" {num}") or f"/{num}" in combined:
            return num

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_norm = normalizar_texto(nombre_producto)
    p_norm = normalizar_texto(presentacion)
    combined_query = normalizar_texto(f"{n_norm} {p_norm}")

    master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")
    master_norm = {normalizar_texto(k): v for k, v in master_dict.items()}

    if combined_query in master_norm:
        return clean_ean_code(master_norm[combined_query])
    if n_norm in master_norm:
        return clean_ean_code(master_norm[n_norm])

    tokens_query = set(re.findall(r'\b[A-Z0-9]+\b', combined_query))
    tokens_query = {t for t in tokens_query if len(t) > 1 and t not in {"ML", "CL", "L", "OZ", "CON", "SIN", "EA", "UN"}}

    if not tokens_query:
        return "S/C"

    mejor_codigo = "S/C"
    max_coincidentes = 0

    for m_key, m_code in master_norm.items():
        tokens_master = set(re.findall(r'\b[A-Z0-9]+\b', m_key))
        comunes = tokens_query.intersection(tokens_master)
        score = len(comunes)
        
        if score > max_coincidentes and score >= 2:
            max_coincidentes = score
            mejor_codigo = clean_ean_code(m_code)

    return mejor_codigo if mejor_codigo != "S/C" else "S/C"

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS Multi-Proveedor</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura (Por Página)", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Sistema Sintaxis Corregida</p>", unsafe_allow_html=True)

if st.sidebar.button("🔄 Reiniciar Historial y Filtros"):
    st.session_state["paginas_procesadas_historial"] = set()
    st.session_state["duplicados_confirmados_sesion"] = set()
    st.session_state["factura_data"] = None
    st.success("Historial limpiado.")
    time.sleep(0.5)
    st.rerun()

# ==========================================
# MÓDULO 1: PROCESAR FACTURA POR PÁGINA
# ==========================================
if menu_opcion == "📄 Procesar Factura (Por Página)":
    st.markdown("<h2>📄 Procesador de Facturas (Multi-Proveedor)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura. El sistema calcula automáticamente los costos unitarios netos por cada unidad/lata.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        st.info("💡 Sube tu factura (PDF multi-página o imagen).")
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu archivo", type=["pdf", "png", "jpg", "jpeg"])
    
    pagina_a_procesar = 1
    if archivo_subido is not None and archivo_subido.name.endswith('.pdf'):
        try:
            import fitz # PyMuPDF
            archivo_subido.seek(0)
            doc_pdf = fitz.open(stream=archivo_subido.read(), filetype="pdf")
            total_paginas = len(doc_pdf)
            st.markdown(f"📄 **PDF detectado con {total_paginas} página(s).**")
            pagina_a_procesar = st.selectbox("Selecciona la página exacta a procesar:", list(range(1, total_paginas + 1)))
        except Exception:
            pagina_a_procesar = st.number_input("Número de página a procesar:", min_value=1, value=1, step=1)

    if archivo_subido is not None:
        if st.button(f"🚀 Procesar Página #{pagina_a_procesar}"):
            with st.spinner(f"🔍 Analizando página #{pagina_a_procesar}..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave de API configurada.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    
                    target_image = None
                    if "pdf" in f_type.lower():
                        import fitz
                        doc_pdf = fitz.open(stream=file_bytes, filetype="pdf")
                        page_obj = doc_pdf[pagina_a_procesar - 1]
                        pix = page_obj.get_pixmap(dpi=150)
                        target_image = Image.open(io.BytesIO(pix.tobytes("jpeg")))
                    else:
                        target_image = Image.open(io.BytesIO(file_bytes))

                    prompt_deteccion = (
                        "Analiza esta página de documento comercial e identifica estrictamente el nombre comercial del proveedor emisor y el número de factura si lo hubiera. "
                        "Devuelve únicamente un JSON con esta estructura: {\"proveedor_detectado\": \"NOMBRE DEL PROVEEDOR\", \"numero_factura\": \"NUMERO\"}"
                    )
                    
                    response_det = model.generate_content([target_image, prompt_deteccion])
                    raw_det_text = response_det.text.strip()
                    if raw_det_text.startswith("```json"): raw_det_text = raw_det_text[7:]
                    if raw_det_text.endswith("```"): raw_det_text = raw_det_text[:-3]
                    
                    det_json = json.loads(raw_det_text.strip())
                    nombre_detectado_raw = str(det_json.get("proveedor_detectado", "PROVEEDOR GENERAL")).upper().strip()
                    num_factura_detectada = str(det_json.get("numero_factura", "S/N")).upper().strip()

                    firma_pagina = f"{archivo_subido.name}_{num_factura_detectada}_PAG_{pagina_a_procesar}"
                    
                    if firma_pagina in st.session_state["paginas_procesadas_historial"]:
                        st.warning(f"⚠️ La página #{pagina_a_procesar} ya fue procesada previamente en esta sesión.")

                    supp_mem = st.session_state["supplier_memory"]
                    if not isinstance(supp_mem, dict): supp_mem = {}
                    
                    prov_encontrado = None
                    for p_key in supp_mem.keys():
                        if p_key in nombre_detectado_raw or nombre_detectado_raw in p_key:
                            prov_encontrado = p_key
                            break
                    
                    if not prov_encontrado:
                        prov_encontrado = nombre_detectado_raw
                        supp_mem[prov_encontrado] = {
                            "nombre": prov_encontrado,
                            "tipo_formato": "factura_desglose_personalizado",
                            "instruccion_prompt": f"Analiza esta página de {prov_encontrado} renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
                        }
                        st.session_state["supplier_memory"] = supp_mem
                        save_json_file(SUPPLIER_MEMORY_FILE, supp_mem)

                    prov_dict_data = supp_mem.get(prov_encontrado, {})
                    if not isinstance(prov_dict_data, dict): prov_dict_data = {}
                    instruccion_proveedor = prov_dict_data.get("instruccion_prompt", "Extrae todos los ítems.")

                    prompt_unificado = (
                        f"Estás procesando la página {pagina_a_procesar} de una factura del proveedor: '{prov_encontrado}'. "
                        f"Instrucción específica de su perfil: {instruccion_proveedor} "
                        "Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario' (precio neto de la caja/paquete sin ITBIS), 'descuento_porcentaje' y 'monto_neto' (subtotal de la línea sin ITBIS). "
                        "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                        '{"paginacion": "' + str(pagina_a_procesar) + '", "proveedor_detectado": "' + prov_encontrado + '", "subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"descripcion": "...", "tamano": "500ML", "cantidad": 1.0, "unidad": "12 PZAS", "precio_unitario": 344.07, "descuento_porcentaje": 0.0, "monto_neto": 3440.70}]}. '
                        "Respuesta JSON pura."
                    )

                    response = model.generate_content([target_image, prompt_unificado])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("
