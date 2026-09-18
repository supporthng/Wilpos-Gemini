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
    page_title="WilPOS - Sistema Integral de Inventario", 
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
# BÚSQUEDA UNIVERSAL DE CLAVES API Y ARCHIVOS
# ------------------------------------------
gemini_paid_candidates = [st.secrets.get("GEMINI_API_KEY_PAID") if "GEMINI_API_KEY_PAID" in st.secrets else None, os.environ.get("GEMINI_API_KEY_PAID")]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else None, st.secrets.get("GOOGLE_API_KEY") if "GOOGLE_API_KEY" in st.secrets else None, os.environ.get("GEMINI_API_KEY"), os.environ.get("GOOGLE_API_KEY")]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

BARCODE_CACHE_FILE = "codigos_escaneados_memoria.json"
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

if "barcode_cache" not in st.session_state: st.session_state["barcode_cache"] = load_json_file(BARCODE_CACHE_FILE, "dict")
if "master_catalog" not in st.session_state: st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE, "dict")
if "master_meta" not in st.session_state: st.session_state["master_meta"] = load_json_file(MASTER_META_FILE, "dict")
if "processing_history" not in st.session_state: st.session_state["processing_history"] = load_json_file(HISTORY_FILE, "list")
if "web_excel_queue" not in st.session_state: st.session_state["web_excel_queue"] = []
if "scanned_buffer" not in st.session_state: st.session_state["scanned_buffer"] = []

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
def save_cache(cache_dict): save_json_file(BARCODE_CACHE_FILE, cache_dict)

def add_to_history(entry):
    hist = st.session_state.get("processing_history", [])
    if not isinstance(hist, list): hist = []
    hist.insert(0, entry)
    st.session_state["processing_history"] = hist
    save_json_file(HISTORY_FILE, hist)

def render_master_status_banner():
    master_dict = st.session_state["master_catalog"]
    supps = st.session_state.get("supplier_memory", {})
    queue_count = len(st.session_state.get("web_excel_queue", []))
    st.success(f"🟢 **WilPOS Reglas de Oro** | Maestro: **{len(master_dict):,}** prods | Proveedores: **{len(supps)}** | En Excel: **{queue_count}**")

def safe_float(val, default=0.0):
    try: return float(val)
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
    if full_query in master_dict: return clean_ean_code(master_dict[full_query])
    if clean_name in master_dict: return clean_ean_code(master_dict[clean_name])
        
    query_words = [w for w in re.findall(r'\w+', full_query.upper()) if len(w) > 2]
    if not query_words: return "S/C"
    
    best_code = "S/C"
    max_matches = 0
    
    for m_name, m_code in master_dict.items():
        m_upper = str(m_name).upper()
        master_words = [w for w in re.findall(r'\w+', m_upper) if len(w) > 2]
        matches = sum(1 for qw in query_words for mw in master_words if qw == mw or (len(qw) >= 4 and (qw in mw or mw in qw)))
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

def online_barcode_lookup_open(barcode_str):
    cache = st.session_state["barcode_cache"]
    if barcode_str in cache:
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        # Uso del modelo estándar oficial y estable
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = (
            f"Busca estrictamente el producto comercial exacto, licor, bebida o artículo asociado "
            f"únicamente al código de barras EAN/UPC: '{barcode_str}'. "
            "Devuelve un JSON puro con el nombre comercial exacto, marca y presentación oficial en mayúsculas: "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si no lo encuentras con certeza, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO EN INTERNET'}."
        )
        
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        
        data = json.loads(txt.strip())
        desc = data.get("descripcion", "PRODUCTO DESCONOCIDO EN INTERNET")
        
        if desc and "DESCONOCIDO" not in desc:
            cache[barcode_str] = desc
            st.session_state["barcode_cache"] = cache
            save_cache(cache)
            return desc, "🌐 (Consultado en Web)"
        else:
            return "PRODUCTO DESCONOCIDO EN INTERNET", "⚠️ (No encontrado)"
            
    except Exception:
        return "PRODUCTO DESCONOCIDO EN INTERNET", "⚠️ (Error técnico)"

def process_invoice_smart_router(file_obj, file_type, use_paid_gemini=False):
    prompt_detect = "Identifica el nombre comercial del proveedor emisor de esta factura (ej: CND, BEES, ALVAREZ & SANCHEZ, GONZALEZ CUESTA). Devuelve un JSON puro: {'proveedor': 'NOMBRE'}"
    
    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini: active_key = ACTIVE_GEMINI_FREE_KEY

    genai.configure(api_key=active_key if active_key else ACTIVE_GEMINI_FREE_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
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
        "Extrae ABSOLUTAMENTE TODOS LOS RENGLONES/PRODUCTOS que aparecen en la factura, sin duplicar ni omitir ninguna. "
        "Para cada renglón extrae en un JSON bajo la clave 'items': "
        "1. 'descripcion': nombre del producto. "
        "2. 'tamano': presentación individual (ej: '750 ML', '70 CL'). "
        "3. 'cantidad': cantidad comprada (ej: 1.0). "
        "4. 'unidad': unidad de empaque (ej: '6 PZA', 'BOT'). "
        "5. 'valor_con_itbis': monto TOTAL INCLUYENDO ITBIS de la línea. "
        "6. 'descuento_monto': monto del descuento (ej: 0.0). "
        "Estructura JSON exacta: "
        '{"proveedor": "' + supplier_detected + '", "items": [{"descripcion": "...", "tamano": "...", "cantidad": 1.0, "unidad": "...", "valor_con_itbis": 0.0, "descuento_monto": 0.0}]}. '
        "Respuesta JSON pura."
    )

    for _ in range(2):
        try:
            file_obj.seek(0)
            response = model.generate_content([image_input, prompt_main])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), f"✅ Éxito ({supplier_detected})"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower(): return None, "QUOTA_EXCEEDED"
            time.sleep(1)
    return None, last_err

# ==========================================
# MENÚ LATERAL DE NAVEGACIÓN
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS Sistema</h3>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio("Navegación", [
    "📄 Factura Individual", 
    "📂 Múltiples Facturas (Lote)", 
    "🌐 Consulta Web de Productos", 
    "📁 Actualizar Catálogo Maestro", 
    "🏢 Perfiles de Proveedores", 
    "📜 Historial de Procesados", 
    "📋 Códigos Almacenados"
])

st.sidebar.markdown("---")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    try:
        st.markdown("<h2>📊 Automatizador <span style='color: #0284c7;'>Reglas de Oro Blindadas</span></h2>", unsafe_allow_html=True)
        render_master_status_banner()

        if "single_processed_data" not in st.session_state: st.session_state["single_processed_data"] = None
        if "single_prov_det" not in st.session_state: st.session_state["single_prov_det"] = "GENERAL"
        if "single_filename" not in st.session_state: st.session_state["single_filename"] = ""

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        c_col1, _ = st.columns([1, 3])
        with c_col1: margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file is not None:
            if st.button("🚀 Procesar Factura"):
                file_type = getattr(uploaded_file, 'type', 'image/jpeg')
                with st.spinner("Procesando factura y calculando costos netos..."):
                    parsed_data, success_msg = process_invoice_smart_router(uploaded_file, file_type, use_paid_gemini=use_gemini_paid_api)

                if success_msg == "QUOTA_EXCEEDED":
                    st.error("⚠️ Límite de cuota alcanzado. Activa 'Gemini Paid'.")
                elif not parsed_data or not isinstance(parsed_data, dict):
                    st.error(f"⚠️ Error al procesar: {success_msg}")
                else:
                    st.session_state["single_processed_data"] = parsed_data
                    st.session_state["single_prov_det"] = parsed_data.get("proveedor", "GENERAL")
                    st.session_state["single_filename"] = uploaded_file.name
                    st.success(f"{success_msg} | Proveedor: **{st.session_state['single_prov_det']}**")

        if st.session_state["single_processed_data"] is not None:
            parsed_data = st.session_state["single_processed_data"]
            prov_det = st.session_state["single_prov_det"]
            rows_preview = []
            calc_subtotal_bruto = calc_total_descuento = calc_subtotal_sin_itbis = calc_total_con_itbis = 0.0

            for idx, item in enumerate(parsed_data.get("items", []), start=1):
                if not isinstance(item, dict): continue
                desc_raw = str(item.get("descripcion") or "").strip()
                if not desc_raw: continue

                unidad_txt = str(item.get("unidad") or "")
                tamano_txt = str(item.get("tamano") or "")
                clean_full_name, clean_pres = clean_product_name_and_presentation(desc_raw, tamano_txt)
                resolved_code = get_flexible_master_barcode(clean_full_name, clean_pres)

                cant_comprada = safe_float(item.get("cantidad") or 1, 1.0)
                val_total_con_itbis_linea = safe_float(item.get("valor_con_itbis") or item.get("importe") or 0)
                descuento_monto_linea = safe_float(item.get("descuento_monto") or 0)
                
                if val_total_con_itbis_linea > 0:
                    val_sin_itbis = val_total_con_itbis_linea / 1.18
                    val_neto_con_itbis = val_total_con_itbis_linea
                    val_bruto_linea = val_neto_con_itbis + descuento_monto_linea
                else:
                    val_bruto_linea = safe_float(item.get("valor_bruto") or 0)
                    val_neto_con_itbis = val_bruto_linea - descuento_monto_linea
                    val_sin_itbis = val_neto_con_itbis / 1.18

                calc_subtotal_bruto += val_bruto_linea
                calc_total_descuento += descuento_monto_linea
                calc_subtotal_sin_itbis += val_sin_itbis
                calc_total_con_itbis += val_neto_con_itbis
                
                empaque_val = parse_empaque_blindado(unidad_txt, clean_full_name, clean_pres)
                total_unidades = int(cant_comprada * empaque_val)
                costo_unitario_real = round(val_sin_itbis / total_unidades, 2) if total_unidades > 0 else 0.0
                if costo_unitario_real <= 0: continue

                raw_pv = (costo_unitario_real * (1 + (margen_ganancia / 100.0))) * 1.18
                precio_venta = round_to_nearest_5(raw_pv)
                
                rows_preview.append({
                    "No.": idx, "Código EAN Maestro": str(resolved_code), "Nombre Producto": clean_full_name,
                    "Presentación": clean_pres if clean_pres else "S/P", "Cant. Compra": cant_comprada,
                    "Unidad": unidad_txt, "Empaque": empaque_val, "Stock Unidades": total_unidades,
                    "Descuento Monto": descuento_monto_linea, "Costo Unitario Real": costo_unitario_real, "Precio Venta": precio_venta
                })

            if rows_preview:
                df_resultado = pd.DataFrame(rows_preview)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                wb = openpyxl.Workbook()
                ws_prod = wb.active
                ws_prod.title = "Productos"
                ws_prod.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                
                for item_dict in rows_preview:
                    code_to_save = str(item_dict["Código EAN Maestro"])
                    desc_val = item_dict["Descuento Monto"]
                    desc_activo = "Sí" if desc_val > 0 else "No"
                    row_idx = ws_prod.max_row + 1
                    ws_prod.append([
                        item_dict["Nombre Producto"], item_dict["Presentación"], str(code_to_save),
                        "General", "producto", round_to_nearest_5(item_dict["Precio Venta"]),
                        round(item_dict["Costo Unitario Real"], 2), int(item_dict["Stock Unidades"]),
                        5, 0.18, "unidad", "No", int(item_dict["Empaque"]), "No", 0, round(desc_val, 2), None, desc_activo, f"Descuento: RD$ {desc_val:,.2f}" if desc_val > 0 else None
                    ])
                    ws_prod.cell(row=row_idx, column=3).number_format = '@'
                    ws_prod.cell(row=row_idx, column=3).value = str(code_to_save)
                
                output = io.BytesIO()
                wb.save(output)
                if st.download_button("📥 Descargar Excel WilPOS", output.getvalue(), f"Inventario_{prov_det.replace('&', 'Y').replace(' ', '_')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
                    add_to_history({
                        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "proveedor": prov_det,
                        "archivo": st.session_state["single_filename"] or "Factura", "total_productos": len(rows_preview),
                        "subtotal": round(calc_subtotal_sin_itbis, 2), "total": round(calc_total_con_itbis, 2)
                    })
    except Exception as e:
        st.error(f"⚠️ Error: {e}")

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por Lotes</h2>", unsafe_allow_html=True)
    render_master_status_banner()
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    uploaded_files = st.file_uploader("📂 Sube tus facturas", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
    st.markdown('</div>', unsafe_allow_html=True)
    if uploaded_files:
        st.info(f"Se cargaron {len(uploaded_files)} archivos para procesamiento en lote.")

# ==========================================
# MÓDULO 3: CONSULTA WEB (ESCANEO EN RÁFAGA + REVISIÓN MANUAL)
# ==========================================
elif modulo == "🌐 Consulta Web de Productos":
    try:
        st.markdown("<h2>🌐 Módulo de Consulta Web <span style='color: #0284c7;'>(Escaneo Libre y Revisión)</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Escanea todos los productos que desees en secuencia. Cada uno aparecerá en una bandeja de revisión para que decidas cuál agregar al Excel.</p>", unsafe_allow_html=True)
        st.markdown("---")

        if "web_excel_queue" not in st.session_state: st.session_state["web_excel_queue"] = []
        if "scanned_buffer" not in st.session_state: st.session_state["scanned_buffer"] = []

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown("### 📌 Lector de Códigos Abierto (Ráfaga)")
        
        def handle_burst_scan():
            val = st.session_state.get("burst_barcode_input", "").strip()
            if val:
                clean_bc = str(val).strip()
                with st.spinner(f"Consultando el código {clean_bc}..."):
                    desc, status = online_barcode_lookup_open(clean_bc)
                
                st.session_state["scanned_buffer"].insert(0, {
                    "codigo": clean_bc,
                    "descripcion": desc,
                    "estado_msg": status,
                    "cantidad": 1
                })
                st.session_state["burst_barcode_input"] = ""

        st.text_input(
            "Escanea o ingresa código de barras (El campo se limpia solo tras cada lectura)", 
            placeholder="Pasa tu lector aquí...", 
            key="burst_barcode_input",
            on_change=handle_burst_scan
        )

        st.markdown('</div>', unsafe_allow_html=True)

        buffer_list = st.session_state.get("scanned_buffer", [])
        if buffer_list:
            st.markdown("### 📦 Bandeja de Productos Escaneados (Pendientes de Agregar)")
            
            for idx, item in enumerate(buffer_list):
                with st.container():
                    cols = st.columns([1.5, 3, 1, 1, 1])
                    with cols[0]:
                        st.markdown(f"**Código:** `{item['codigo']}`")
                        st.caption(item['estado_msg'])
                    with cols[1]:
                        edited_desc = st.text_input("Descripción", value=item['descripcion'], key=f"buf_desc_{idx}")
                    with cols[2]:
                        edited_qty = st.number_input("Cantidad", min_value=1, value=item['cantidad'], key=f"buf_qty_{idx}")
                    with cols[3]:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("➕ Agregar", key=f"btn_add_{idx}"):
                            new_entry = {
                                "Código de Barra": str(item['codigo']),
                                "Descripción": str(edited_desc).strip().upper(),
                                "Cantidad": int(edited_qty)
                            }
                            st.session_state["web_excel_queue"].append(new_entry)
                            st.session_state["scanned_buffer"].pop(idx)
                            st.success("¡Agregado al Excel!")
                            st.rerun()
                    with cols[4]:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("❌ Descartar", key=f"btn_drop_{idx}"):
                            st.session_state["scanned_buffer"].pop(idx)
                            st.rerun()
                st.markdown("---")

        st.markdown("### 📊 Vista Previa en Vivo del Excel Final")
        queue = st.session_state.get("web_excel_queue", [])
        
        if queue:
            df_web = pd.DataFrame(queue)
            df_web["Código de Barra"] = df_web["Código de Barra"].astype(str)
            st.dataframe(df_web, use_container_width=True, hide_index=True)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Consulta Web"
            ws.append(["Código de Barra", "Descripción", "Cantidad"])

            for idx_row, row_item in enumerate(queue, start=2):
                ws.append([str(row_item["Código de Barra"]), row_item["Descripción"], int(row_item["Cantidad"])])
                ws.cell(row=idx_row, column=1).number_format = '@'

            buffer_excel = io.BytesIO()
            wb.save(buffer_excel)

            col_act1, col_act2 = st.columns(2)
            with col_act1:
                st.download_button(
                    label="📥 Descargar Excel de Consulta Web",
                    data=buffer_excel.getvalue(),
                    file_name=f"Consulta_Web_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_act2:
                if st.button("🗑️ Limpiar Todo el Excel"):
                    st.session_state["web_excel_queue"] = []
                    st.rerun()
        else:
            st.info("ℹ️ No hay elementos en el Excel final todavía. Revisa y agrega los productos desde tu bandeja superior.")
    except Exception as e:
        st.error(f"⚠️ Error: {e}")

# ==========================================
# MÓDULO 4: ACTUALIZAR CATÁLOGO MAESTRO
# ==========================================
elif modulo == "📁 Actualizar Catálogo Maestro":
    st.markdown("<h2>📁 Catálogo Maestro</h2>", unsafe_allow_html=True)
    render_master_status_banner()
    st.markdown("---")
    master_file = st.file_uploader("📂 Sube tu Catálogo Maestro (Excel)", type=["xlsx"])
    if master_file is not None:
        df_master = pd.read_excel(master_file, dtype=str)
        cols = df_master.columns.tolist()
        col_name = st.selectbox("Columna con Nombre", cols)
        col_code = st.selectbox("Columna con Código EAN", cols)
        if st.button("🔄 Guardar Maestro EAN"):
            count = 0
            temp_dict = {}
            for _, row in df_master.iterrows():
                p_name, _ = clean_product_name_and_presentation(row[col_name])
                p_code = clean_ean_code(row[col_code])
                if p_name and p_code != "S/C":
                    temp_dict[p_name] = str(p_code)
                    count += 1
            st.session_state["master_catalog"] = temp_dict
            save_master_to_file()
            save_meta_to_file(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), count)
            st.success(f"¡Catálogo actualizado con {count} productos!")
            st.rerun()

# ==========================================
# MÓDULO 5: PERFILES DE PROVEEDORES
# ==========================================
elif modulo == "🏢 Perfiles de Proveedores":
    st.markdown("<h2>🏢 Memoria de Proveedores</h2>", unsafe_allow_html=True)
    st.markdown("---")
    supps = st.session_state["supplier_memory"]
    for s_name, s_data in supps.items():
        with st.expander(f"🏢 {s_name}"):
            st.write(f"**Formato:** {s_data.get('formato_empaque', 'N/A')}")

# ==========================================
# MÓDULO 6: HISTORIAL DE PROCESADOS
# ==========================================
elif modulo == "📜 Historial de Procesados":
    st.markdown("<h2>📜 Historial de Documentos</h2>", unsafe_allow_html=True)
    st.markdown("---")
    history_list = st.session_state.get("processing_history", [])
    if history_list:
        st.dataframe(pd.DataFrame(history_list), use_container_width=True, hide_index=True)
        if st.button("🗑️ Limpiar Historial"):
            st.session_state["processing_history"] = []
            save_json_file(HISTORY_FILE, [])
            st.rerun()
    else:
        st.info("No hay historial todavía.")

# ==========================================
# MÓDULO 7: CÓDIGOS ALMACENADOS
# ==========================================
elif modulo == "📋 Códigos Almacenados":
    st.markdown("<h2>📋 Memoria de Códigos EAN</h2>", unsafe_allow_html=True)
    render_master_status_banner()
    st.markdown("---")
    master = st.session_state["master_catalog"]
    if master:
        st.dataframe(pd.DataFrame([{"Producto": k, "Código EAN": str(v)} for k, v in master.items()]), use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ No hay productos en el maestro.")
