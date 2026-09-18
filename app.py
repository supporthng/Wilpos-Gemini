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

def render_master_status_banner():
    master_dict = st.session_state["master_catalog"]
    meta = st.session_state.get("master_meta", {})
    supps = st.session_state.get("supplier_memory", {})
    queue_count = len(st.session_state.get("live_excel_queue", []))
    st.success(f"🟢 **WilPOS Reglas de Oro Blindadas** | Maestro: **{len(master_dict):,}** prods | 🏢 Proveedores: **{len(supps)}** | 📋 Cola Excel En Vivo: **{queue_count}**")

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

def online_barcode_lookup(barcode_str):
    """Consulta inteligente en internet usando Gemini con búsqueda web integrada para códigos desconocidos."""
    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-3.6-flash')
        prompt = (
            f"Busca en internet el producto exacto asociado al código de barras EAN: '{barcode_str}'. "
            "Devuelve un JSON puro con el nombre comercial y presentación oficial en mayúsculas: "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si no lo encuentras, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO'}."
        )
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        data = json.loads(txt.strip())
        return data.get("descripcion", "")
    except Exception:
        return ""

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
# MÓDULO 3: CONSULTA NUEVOS PRODUCTOS CON BÚSQUEDA WEB Y CANTIDAD
# ==========================================
if modulo == "🔍 Consulta Nuevos Productos":
    try:
        st.markdown("<h2>🔍 Consulta de Nuevos Productos <span style='color: #0284c7;'>(Búsqueda Web Automática)</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Ingresa el código de barra. Si no está en tu maestro, el sistema lo buscará en internet, te pedirá la cantidad y armará tu Excel en vivo.</p>", unsafe_allow_html=True)
        st.markdown("---")
        render_master_status_banner()

        if "lookup_barcode_input" not in st.session_state: st.session_state["lookup_barcode_input"] = ""
        if "lookup_desc_result" not in st.session_state: st.session_state["lookup_desc_result"] = ""

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown("### 📌 Búsqueda por Código de Barra")
        
        col_s1, col_s2 = st.columns([2, 1])
        with col_s1:
            input_barcode = st.text_input("Código de Barra (EAN)", placeholder="Ej: 080480172022", key="input_bc_widget")
        with col_s2:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_buscar = st.button("🌐 Buscar en Maestro e Internet")

        if btn_buscar and input_barcode.strip():
            bc = clean_ean_code(input_barcode)
            # 1. Buscar en Maestro local
            master_dict = st.session_state.get("master_catalog", {})
            found_name = ""
            for name, code in master_dict.items():
                if str(code).strip() == bc:
                    found_name = name
                    break
            
            # 2. Si no está en maestro, buscar en internet
            if not found_name:
                with st.spinner(f"Buscando el código {bc} en internet..."):
                    found_name = online_barcode_lookup(bc)

            st.session_state["lookup_barcode_input"] = bc
            st.session_state["lookup_desc_result"] = found_name if found_name else "PRODUCTO NO IDENTIFICADO"

        # Mostrar campos de resultado y cantidad
        current_bc = st.session_state.get("lookup_barcode_input", "")
        current_desc = st.session_state.get("lookup_desc_result", "")

        if current_bc:
            st.markdown("---")
            st.markdown("### 📝 Detalle del Producto Encontrado")
            
            col_d1, col_d2 = st.columns([2, 1])
            with col_d1:
                final_desc_input = st.text_input("Descripción Oficial (Nombre + Presentación)", value=current_desc)
            with col_d2:
                final_cant_input = st.number_input("Cantidad Comprada", min_value=1, max_value=10000, value=1, step=1)

            if st.button("➕ Agregar a la Vista Previa del Excel"):
                if final_desc_input.strip() and final_desc_input != "PRODUCTO NO IDENTIFICADO":
                    clean_name, _ = clean_product_name_and_presentation(final_desc_input)
                    new_item = {
                        "Código de Barra": current_bc,
                        "Descripción": clean_name,
                        "Cantidad": int(final_cant_input)
                    }
                    st.session_state["live_excel_queue"].append(new_item)
                    
                    # También guardar opcionalmente en el maestro para futuras ocasiones
                    st.session_state["master_catalog"][clean_name] = current_bc
                    save_master_to_file()

                    st.success(f"✅ ¡Agregado exitosamente: **{clean_name}** (Cant: {final_cant_input}) y guardado en el maestro!")
                    st.session_state["lookup_barcode_input"] = ""
                    st.session_state["lookup_desc_result"] = ""
                    st.rerun()
                else:
                    st.error("Por favor verifica que la descripción sea válida antes de agregarla.")

        st.markdown('</div>', unsafe_allow_html=True)

        # VISTA PREVIA EN VIVO DEL EXCEL
        st.markdown("### 📊 Vista Previa en Vivo del Excel a Generar")
        queue = st.session_state.get("live_excel_queue", [])
        
        if queue:
            df_preview = pd.DataFrame(queue)
            df_preview["Código de Barra"] = df_preview["Código de Barra"].astype(str)
            st.dataframe(df_preview, use_container_width=True, hide_index=True)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Nuevos Productos"
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
                    label="📥 Descargar Excel de Nuevos Productos",
                    data=excel_out.getvalue(),
                    file_name=f"Excel_Nuevos_Productos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_act2:
                if st.button("🗑️ Vaciar Vista Previa"):
                    st.session_state["live_excel_queue"] = []
                    st.rerun()
        else:
            st.info("ℹ️ Tu lista en vivo está vacía. Ingresa un código arriba para buscarlo en internet y agregarlo.")

    except Exception as e:
        st.error("⚠️ Error en Consulta de Nuevos Productos:")
        st.exception(e)

# ==========================================
# MÓDULOS RESTANTES
# ==========================================
elif modulo == "📄 Factura Individual":
    st.info("Módulo activo de Factura Individual.")
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.info("Módulo activo de Lotes.")
elif modulo == "📁 Actualizar Catálogo Maestro":
    st.info("Módulo activo de Catálogo Maestro.")
elif modulo == "🏢 Perfiles de Proveedores":
    st.info("Módulo activo de Proveedores.")
elif modulo == "📜 Historial de Procesados":
    st.info("Módulo activo de Historial.")
elif modulo == "📋 Códigos Almacenados":
    st.info("Módulo activo de Códigos Almacenados.")
