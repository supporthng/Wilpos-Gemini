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
            "CND": {"nombre": "CND", "formato_empaque": "cnd_bees_extractor"},
            "BEES": {"nombre": "BEES", "formato_empaque": "cnd_bees_extractor"}
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

def get_strict_ean_code_and_name(description, invoice_ean=""):
    desc_clean = str(description).strip().upper()
    master = st.session_state["master_catalog"]
    memory = st.session_state["barcode_memory"]
    
    if desc_clean in master: return desc_clean, clean_ean_code(master[desc_clean])
    if desc_clean in memory: return desc_clean, clean_ean_code(memory[desc_clean])
    
    cleaned_invoice_ean = clean_ean_code(invoice_ean)
    if cleaned_invoice_ean != "S/C (Sin Codigo)": return desc_clean, cleaned_invoice_ean
        
    master_keys = list(master.keys())
    if master_keys:
        coincidencias = difflib.get_close_matches(desc_clean, master_keys, n=1, cutoff=0.60)
        if coincidencias: return desc_clean, clean_ean_code(master[coincidencias[0]])
            
    return desc_clean, cleaned_invoice_ean

# Regla estricta: Si es unidad suelta (UN) o no hay empaque por cajas explícito en el texto, el empaque es 1.
def parse_empaque_isolated(supplier_name, tamano_txt="", unidad_txt="", descripcion_txt=""):
    s_name = str(supplier_name).upper()
    u = str(unidad_txt).strip().upper()
    t = str(tamano_txt).strip().upper()
    d = str(descripcion_txt).strip().upper()
    combined = f"{t} {u} {d}"

    # Si la unidad es explícitamente UN (unidad suelta), empaque es 1 sin discusión
    if "UN" in u:
        return 1

    # Buscar patrón explícito de cajas/huacales en la descripción (ej: '24/12OZ')
    match_explicit_slash = re.search(r'\b(24|12|6|48)\s*/', combined)
    if match_explicit_slash:
        return int(match_explicit_slash.group(1))

    if "ALVAREZ" in s_name or "SANCHEZ" in s_name:
        match_4x6 = re.search(r'(\d+)\s*X\s*(\d+)', combined)
        if match_4x6: return int(match_4x6.group(1)) * int(match_4x6.group(2))
    elif "GONZALEZ" in s_name or "CUESTA" in s_name:
        if "CAJ" in u or "CAJA" in u:
            if "12" in combined: return 12
            if "6" in combined: return 6
            if "24" in combined: return 24
            if "48" in combined: return 48
    elif "DEPOT" in s_name or "PRICESMART" in s_name:
        return 1

    match_pza = re.search(r'(\d+)\s*PZA', combined)
    if match_pza: return int(match_pza.group(1))

    # Por defecto, si viene como unidad (UN), respetamos la cantidad exacta con empaque 1
    return 1

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Sistema con Empaque Estricto</p>", unsafe_allow_html=True)
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
        supp_mem[supplier_detected] = {"nombre": supplier_detected, "formato_empaque": "adaptativo", "descripcion": "Registrado automaticamente."}
        save_supplier_memory()

    prompt_main = (
        f"Analiza este documento de compra del proveedor '{supplier_detected}' con absoluta precisión. "
        "REGLA DE ORO: Extrae rigurosamente la unidad de medida exactamente como aparece en la factura (por ejemplo 'UN' para unidades sueltas, 'PC' para bultos/huacales). "
        "Para cada renglón extrae en un JSON bajo la clave 'items': "
        "1. 'codigo_ean': código de barras o código de artículo. "
        "2. 'descripcion': nombre exacto del producto. "
        "3. 'cantidad': cantidad comprada (ej: 47.0). "
        "4. 'tamano': tamaño o presentación si aplica. "
        "5. 'unidad': unidad de medida exacta impresa en la línea (ej: 'UN', 'PC'). "
        "6. 'valor': monto total neto de la línea. "
        "Estructura JSON exacta: "
        '{"proveedor": "' + supplier_detected + '", "items": [{"codigo_ean": "...", "descripcion": "...", "cantidad": 1.0, "tamano": "...", "unidad": "...", "valor": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
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
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    try:
        st.markdown("<h2>📊 Automatizador con <span style='color: #0284c7;'>Respeto Absoluto de Unidades</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Las unidades (UN) se calculan 1 a 1 con su costo exacto.</p>", unsafe_allow_html=True)
        st.markdown("---")
        render_master_status_banner()

        if "single_processed_data" not in st.session_state: st.session_state["single_processed_data"] = None
        if "single_prov_det" not in st.session_state: st.session_state["single_prov_det"] = "GENERAL"
        if "single_filename" not in st.session_state: st.session_state["single_filename"] = ""

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        c_col1, _ = st.columns([1, 3])
        with c_col1: margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        uploaded_file = st.file_uploader("📂 Sube tu factura o tiquete (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file is not None:
            if st.button("🚀 Procesar y Guardar en Historial"):
                file_type = getattr(uploaded_file, 'type', 'image/jpeg')
                with st.spinner("Procesando documento y respetando unidades..."):
                    parsed_data, success_msg = process_invoice_smart_router(uploaded_file, file_type, use_paid_gemini=use_gemini_paid_api)

                if success_msg == "QUOTA_EXCEEDED":
                    st.error("⚠️ Límite de cuota alcanzado. Activa 'Gemini Paid'.")
                elif not parsed_data or not isinstance(parsed_data, dict):
                    st.error(f"⚠️ Error al procesar: {success_msg}")
                else:
                    st.session_state["single_processed_data"] = parsed_data
                    st.session_state["single_prov_det"] = parsed_data.get("proveedor", "GENERAL")
                    st.session_state["single_filename"] = uploaded_file.name
                    st.success(f"{success_msg} | 🏢 Proveedor detectado: **{st.session_state['single_prov_det']}**")

        if st.session_state["single_processed_data"] is not None:
            parsed_data = st.session_state["single_processed_data"]
            prov_det = st.session_state["single_prov_det"]
            
            data_items = parsed_data.get("items", [])
            rows_preview = []
            multiplicador_ganancia = 1 + (margen_ganancia / 100.0)
            calc_subtotal = 0.0

            assigned_barcodes = set()

            for idx, item in enumerate(data_items, start=1):
                if not isinstance(item, dict): continue
                desc_raw = str(item.get("descripcion") or "").strip()
                if not desc_raw: continue

                invoice_ean = str(item.get("codigo_ean") or "")
                resolved_name, resolved_code = get_strict_ean_code_and_name(desc_raw, invoice_ean)

                if resolved_code in assigned_barcodes and resolved_code != "S/C (Sin Codigo)":
                    resolved_code = "S/C (Duplicado - Revisar)"

                if resolved_code != "S/C (Sin Codigo)" and "Duplicado" not in resolved_code:
                    assigned_barcodes.add(resolved_code)

                cant_comprada = safe_float(item.get("cantidad") or 1, 1.0)
                val_neto_linea = safe_float(item.get("valor") or 0)
                tamano_txt = str(item.get("tamano") or "")
                unidad_txt = str(item.get("unidad") or "")
                
                empaque_val = parse_empaque_isolated(prov_det, tamano_txt, unidad_txt, resolved_name)
                total_unidades = int(cant_comprada * empaque_val)

                costo_unitario_real = round(val_neto_linea / total_unidades, 2) if total_unidades > 0 else 0.0
                if costo_unitario_real <= 0: continue

                calc_subtotal += val_neto_linea
                raw_pv = (costo_unitario_real * multiplicador_ganancia) * 1.18
                precio_venta = round_to_nearest_5(raw_pv)
                
                rows_preview.append({
                    "No.": idx,
                    "Código EAN Único": str(resolved_code),
                    "Nombre Producto": resolved_name,
                    "Cant. Compra": cant_comprada,
                    "Unidad": unidad_txt,
                    "Empaque": empaque_val,
                    "Stock Unidades": total_unidades,
                    "Costo Unitario Real": costo_unitario_real,
                    "Precio Venta": precio_venta
                })

            calc_itbis = calc_subtotal * 0.18
            calc_total_factura = calc_subtotal + calc_itbis

            st.markdown("### 📑 Totales del Documento")
            t1, t2, t3 = st.columns(3)
            t1.metric("Subtotal Factura", f"RD$ {calc_subtotal:,.2f}")
            t2.metric("ITBIS Total (18%)", f"RD$ {calc_itbis:,.2f}")
            t3.metric("Total Neto", f"RD$ {calc_total_factura:,.2f}")
            st.markdown("---")

            if rows_preview:
                st.markdown("### ✅ Artículos Procesados Correctamente")
                df_resultado = pd.DataFrame(rows_preview)
                df_resultado["Código EAN Único"] = df_resultado["Código EAN Único"].astype(str)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                wb = openpyxl.Workbook()
                ws_prod = wb.active
                ws_prod.title = "Productos"
                ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                
                for item_dict in rows_preview:
                    code_to_save = item_dict["Código EAN Único"]
                    if "Duplicado" in code_to_save or "Sin Codigo" in code_to_save: code_to_save = "S/C"
                    ws_prod.append([
                        item_dict["Nombre Producto"], str(code_to_save),
                        "General", "producto", item_dict["Precio Venta"],
                        item_dict["Costo Unitario Real"], item_dict["Stock Unidades"],
                        5, 0.18, "unidad", "No", item_dict["Empaque"], "No", 0, 0, None, "No", None
                    ])
                    ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                
                output = io.BytesIO()
                wb.save(output)
                
                if st.download_button("📥 Descargar Excel WilPOS Oficial", output.getvalue(), f"Inventario_{prov_det.replace(' ', '_')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
                    history_entry = {
                        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "proveedor": prov_det,
                        "archivo": st.session_state["single_filename"] or "Factura Individual",
                        "subtotal": round(calc_subtotal, 2),
                        "itbis": round(calc_itbis, 2),
                        "total": round(calc_total_factura, 2),
                        "total_items": len(rows_preview)
                    }
                    add_to_history(history_entry)
    except Exception as e:
        st.error("⚠️ Error en Factura Individual:")
        st.exception(e)

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    try:
        st.markdown("<h2>📂 Procesador por Lotes <span style='color: #0284c7;'>(Protegido)</span></h2>", unsafe_allow_html=True)
        st.markdown("---")
        render_master_status_banner()
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        l_col1, _ = st.columns([1, 3])
        with l_col1: margen_ganancia_lote = st.number_input("⚙️ Ganancia (%) Lote", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        uploaded_files = st.file_uploader("📂 Sube tus facturas (Selección múltiple)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_files:
            if "cached_uploaded_files" not in st.session_state or len(st.session_state["cached_uploaded_files"]) != len(uploaded_files):
                st.session_state["cached_uploaded_files"] = [{"name": f.name, "type": getattr(f, "type", "image/jpeg"), "bytes": f.read()} for f in uploaded_files]

        if "cached_uploaded_files" in st.session_state and st.session_state["cached_uploaded_files"]:
            cached_files = st.session_state["cached_uploaded_files"]
            total_files = len(cached_files)

            if "batch_accumulated_items" not in st.session_state:
                st.session_state["batch_accumulated_items"] = []
                st.session_state["batch_processed_count"] = 0
                st.session_state["is_live_processing"] = False

            processed_so_far = st.session_state["batch_processed_count"]
            b_col1, b_col2 = st.columns(2)
            if b_col1.button("🚀 Iniciar Lote Seguro", type="primary"):
                st.session_state["is_live_processing"] = True
                st.rerun()
            if b_col2.button("🔄 Reiniciar"):
                st.session_state["batch_accumulated_items"] = []
                st.session_state["batch_processed_count"] = 0
                st.session_state["is_live_processing"] = False
                if "cached_uploaded_files" in st.session_state: del st.session_state["cached_uploaded_files"]
                st.rerun()

            if st.session_state.get("is_live_processing", False):
                if processed_so_far < total_files:
                    file_info = cached_files[processed_so_far]
                    st.info(f"⚡ Procesando {processed_so_far + 1} de {total_files}: `{file_info['name']}`...")
                    st.progress(processed_so_far / total_files)

                    parsed_data, _ = process_invoice_smart_router(io.BytesIO(file_info["bytes"]), file_info["type"], use_paid_gemini=use_gemini_paid_api)
                    time.sleep(1.0)

                    if parsed_data and isinstance(parsed_data, dict):
                        prov_det = parsed_data.get("proveedor", "GENERAL")
                        for itm in parsed_data.get("items", []):
                            if isinstance(itm, dict) and str(itm.get("descripcion") or "").strip():
                                itm["_prov"] = prov_det
                                st.session_state["batch_accumulated_items"].append(itm)
                    st.session_state["batch_processed_count"] += 1
                    st.rerun()
                else:
                    st.session_state["is_live_processing"] = False
                    st.rerun()

            if st.session_state["batch_processed_count"] > 0:
                st.markdown("---")
                raw_items = st.session_state["batch_accumulated_items"]
                multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)
                processed_rows = []
                batch_assigned_barcodes = set()

                for item in raw_items:
                    desc_raw = str(item.get("descripcion") or "").strip()
                    if not desc_raw: continue
                    invoice_ean = str(item.get("codigo_ean") or "")
                    prov_det = str(item.get("_prov") or "GENERAL")
                    resolved_name, resolved_code = get_strict_ean_code_and_name(desc_raw, invoice_ean)
                    
                    if resolved_code in batch_assigned_barcodes and resolved_code != "S/C (Sin Codigo)":
                        resolved_code = "S/C"
                    if resolved_code != "S/C (Sin Codigo)":
                        batch_assigned_barcodes.add(resolved_code)

                    cant_comprada = safe_float(item.get("cantidad") or 1, 1.0)
                    val_neto_linea = safe_float(item.get("valor") or 0)
                    tamano_txt = str(item.get("tamano") or "")
                    unidad_txt = str(item.get("unidad") or "")
                    
                    empaque_val = parse_empaque_isolated(prov_det, tamano_txt, unidad_txt, resolved_name)
                    total_unidades = int(cant_comprada * empaque_val)
                    costo_unitario_real = round(val_neto_linea / total_unidades, 2) if total_unidades > 0 else 0.0
                    if costo_unitario_real <= 0: continue

                    raw_pv = (costo_unitario_real * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)

                    processed_rows.append({
                        "Nombre": resolved_name, "Código Barra": str(resolved_code),
                        "Categoría": "General", "Tipo": "producto", "Precio Venta": precio_venta,
                        "Costo": costo_unitario_real, "Stock": total_unidades, "Stock Mínimo": 5,
                        "ITBIS": 0.18, "Unidad Medida": "unidad", "Venta Granel": "No",
                        "Cantidad Empaque": empaque_val, "Precio Variable": "No",
                        "Descuento %": 0, "Descuento Monto": 0, "Precio Especial": None,
                        "Descuento Activo": "No", "Descuento Nota": None
                    })

                if processed_rows:
                    df_temp = pd.DataFrame(processed_rows)
                    df_grouped = df_temp.groupby(['Código Barra', 'Nombre'], as_index=False).agg({
                        'Stock': 'sum', 'Costo': 'mean', 'Precio Venta': 'mean',
                        'Categoría': 'first', 'Tipo': 'first', 'Stock Mínimo': 'first',
                        'ITBIS': 'first', 'Unidad Medida': 'first', 'Venta Granel': 'first',
                        'Cantidad Empaque': 'first', 'Precio Variable': 'first',
                        'Descuento %': 'first', 'Descuento Monto': 'first',
                        'Precio Especial': 'first', 'Descuento Activo': 'first', 'Descuento Nota': 'first'
                    })
                    st.dataframe(df_grouped, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error("⚠️ Error en Lotes:")
        st.exception(e)

# ==========================================
# MÓDULO 3: ACTUALIZAR CATÁLOGO MAESTRO
# ==========================================
elif modulo == "📁 Actualizar Catálogo Maestro":
    try:
        st.markdown("<h2>📁 Catálogo Maestro</h2>", unsafe_allow_html=True)
        render_master_status_banner()
        st.markdown("---")
        master_file = st.file_uploader("📂 Sube tu Catálogo Maestro (Excel)", type=["xlsx"])
        if master_file is not None:
            df_master = pd.read_excel(master_file)
            cols = df_master.columns.tolist()
            col_name = st.selectbox("Columna con Nombre", cols)
            col_code = st.selectbox("Columna con Código EAN", cols)
            if st.button("🔄 Guardar Maestro EAN"):
                count = 0
                temp_dict = {}
                for _, row in df_master.iterrows():
                    p_name = str(row[col_name]).strip().upper()
                    p_code = clean_ean_code(row[col_code])
                    if p_name and p_code != "S/C (Sin Codigo)":
                        temp_dict[p_name] = p_code
                        count += 1
                st.session_state["master_catalog"] = temp_dict
                save_master_to_file()
                save_meta_to_file(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), count)
                st.success(f"¡Catálogo EAN actualizado con {count} productos!")
                st.rerun()
    except Exception as e:
        st.error("⚠️ Error en Maestro:")
        st.exception(e)

# ==========================================
# MÓDULO 4: PERFILES DE PROVEEDORES
# ==========================================
elif modulo == "🏢 Perfiles de Proveedores":
    try:
        st.markdown("<h2>🏢 Memoria de Proveedores</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Perfiles y reglas aisladas por proveedor.</p>", unsafe_allow_html=True)
        st.markdown("---")
        supps = st.session_state["supplier_memory"]
        if supps:
            for s_name, s_data in supps.items():
                with st.expander(f"🏢 {s_name}"):
                    st.write(f"**Formato / Regla:** {s_data.get('formato_empaque', 'N/A')}")
                    st.write(f"**Descripcion:** {s_data.get('descripcion', 'N/A')}")
        else:
            st.info("No hay proveedores en memoria aún.")
    except Exception as e:
        st.error("⚠️ Error:")
        st.exception(e)

# ==========================================
# MÓDULO 5: HISTORIAL DE PROCESADOS
# ==========================================
elif modulo == "📜 Historial de Procesados":
    try:
        st.markdown("<h2>📜 Historial de Documentos Procesados</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Registro cronológico de todas las facturas y tiquetes procesados.</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        history_list = st.session_state.get("processing_history", [])
        if not isinstance(history_list, list): history_list = []
            
        if history_list:
            df_hist = pd.DataFrame(history_list)
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
            
            if st.button("🗑️ Limpiar Historial"):
                st.session_state["processing_history"] = []
                save_json_file(HISTORY_FILE, [])
                st.success("¡Historial limpiado exitosamente!")
                st.rerun()
        else:
            st.info("No hay documentos procesados en el historial todavía.")
    except Exception as e:
        st.error("⚠️ Error en Historial:")
        st.exception(e)

# ==========================================
# MÓDULO 6: CÓDIGOS ALMACENADOS
# ==========================================
elif modulo == "📋 Códigos Almacenados":
    try:
        st.markdown("<h2>📋 Memoria de Códigos EAN</h2>", unsafe_allow_html=True)
        render_master_status_banner()
        st.markdown("---")
        master = st.session_state["master_catalog"]
        if master:
            st.dataframe(pd.DataFrame([{"Producto": k, "Código EAN": v} for k, v in master.items()]), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ No hay productos cargados en el maestro.")
    except Exception as e:
        st.error("⚠️ Error:")
        st.exception(e)
