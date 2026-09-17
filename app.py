import io
import json
import os
import time
from datetime import datetime
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
    page_title="WilPOS - Automatizador Inteligente con Memoria de Proveedores", 
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
    [data-testid="stMetricValue"] {
        font-size: 1.25rem !important;
        font-weight: 700;
        white-space: nowrap;
        overflow: visible;
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
    st.secrets.get("GEMINI_API_KEY_PAID") if "GEMINI_API_KEY_PAID" in st.secrets else None,
    os.environ.get("GEMINI_API_KEY_PAID")
]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [
    st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else None,
    st.secrets.get("GOOGLE_API_KEY") if "GOOGLE_API_KEY" in st.secrets else None,
    st.secrets.get("GEMINI_API_KEY_1") if "GEMINI_API_KEY_1" in st.secrets else None,
    os.environ.get("GEMINI_API_KEY"),
    os.environ.get("GOOGLE_API_KEY"),
    os.environ.get("GEMINI_API_KEY_1")
]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

openai_key_candidates = [
    st.secrets.get("OPENAI_API_KEY") if "OPENAI_API_KEY" in st.secrets else None,
    os.environ.get("OPENAI_API_KEY")
]
ACTIVE_OPENAI_KEY = next((k for k in openai_key_candidates if k and str(k).strip()), None)

BARCODE_MEMORY_FILE = "codigos_escaneados_memoria.json"
MASTER_CATALOG_FILE = "catalogo_maestro_sistema.json"
MASTER_META_FILE = "catalogo_maestro_meta.json"
SUPPLIER_MEMORY_FILE = "proveedores_formatos_memoria.json"

def load_json_file(filepath):
    data = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    return data

def save_json_file(filepath, data):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE)

if "master_meta" not in st.session_state:
    st.session_state["master_meta"] = load_json_file(MASTER_META_FILE)

# Inicialización y Persistencia de la Memoria de Proveedores
if "supplier_memory" not in st.session_state:
    loaded_suppliers = load_json_file(SUPPLIER_MEMORY_FILE)
    if not loaded_suppliers:
        loaded_suppliers = {
            "PRICESMART": {
                "nombre": "PRICESMART",
                "formato_codigo": "codigo_item_numerico",
                "regla_empaque": "Unidades sueltas o empaques indicados en descripción/tamaño.",
                "descripcion": "Facturas estilo tiquet, con códigos numéricos de ítem a la izquierda y descripciones literales sin alteración."
            },
            "GONZALEZ CUESTA": {
                "nombre": "GONZALEZ CUESTA",
                "formato_codigo": "ean_13_digitos_abajo_sap",
                "regla_empaque": "Multiplicación obligatoria de Cajas (CAJ) por unidades de la UMV (ej: CAJ / 12 PZA = x12).",
                "descripcion": "Facturas de crédito fiscal con código SAP corto arriba y código de barras EAN largo de 13 dígitos abajo."
            },
            "AD ROYAL LICOR": {
                "nombre": "AD ROYAL LICOR",
                "formato_codigo": "ean_13_digitos_abajo_sap",
                "regla_empaque": "Multiplicación obligatoria de Cajas (CAJ) por unidades de la UMV (ej: CAJ / 12 PZA = x12).",
                "descripcion": "Facturas de crédito fiscal formales con doble código (SAP / EAN) en columna izquierda."
            }
        }
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_suppliers)
    st.session_state["supplier_memory"] = loaded_suppliers

def save_master_to_file():
    save_json_file(MASTER_CATALOG_FILE, st.session_state["master_catalog"])

def save_meta_to_file(timestamp_str, count):
    meta = {"ultima_actualizacion": timestamp_str, "total_productos": count}
    st.session_state["master_meta"] = meta
    save_json_file(MASTER_META_FILE, meta)

def save_supplier_memory():
    save_json_file(SUPPLIER_MEMORY_FILE, st.session_state["supplier_memory"])

def render_master_status_banner():
    master_dict = st.session_state["master_catalog"]
    meta = st.session_state.get("master_meta", {})
    suppliers = st.session_state.get("supplier_memory", {})
    total_prod = len(master_dict)
    total_supp = len(suppliers)
    ultima_act = meta.get("ultima_actualizacion", "Desconocida")

    st.success(f"🟢 **WilPOS Activo** | Catálogo: **{total_prod:,}** prods | 🏢 Proveedores Registrados: **{total_supp}** | 🕒 Última act: **{ultima_act}**")

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

def clean_code(code_val):
    if not code_val:
        return "S/C (Sin Código)"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'):
        s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]:
        return "S/C (Sin Código)"
    return s_val

def get_resolved_code_and_name(description, invoice_code=""):
    if invoice_code and str(invoice_code).strip() not in ["", "nan", "None", "S/C"]:
        return str(description).strip().upper(), clean_code(invoice_code)

    desc_clean = str(description).strip().upper()
    master = st.session_state["master_catalog"]
    memory = st.session_state["barcode_memory"]
    
    if desc_clean in master:
        return desc_clean, clean_code(master[desc_clean])
    if desc_clean in memory:
        return desc_clean, clean_code(memory[desc_clean])
        
    master_keys = list(master.keys())
    if master_keys:
        coincidencias = difflib.get_close_matches(desc_clean, master_keys, n=1, cutoff=0.80)
        if coincidencias:
            matched_name = coincidencias[0]
            return desc_clean, clean_code(master[matched_name])
            
    return desc_clean, "S/C (Sin Código)"

def parse_empaque_by_supplier(supplier_name, umv_txt, descripcion_txt=""):
    u = str(umv_txt).strip().upper()
    d = str(descripcion_txt).strip().upper()
    combined = d + " " + u
    
    # Proveedores formales con cajas explícitas (González Cuesta / Ad Royal Licor)
    if any(k in supplier_name for k in ["GONZALEZ", "ROYAL", "CUESTA"]):
        match_pza = re.search(r'(\d+)\s*PZA', combined)
        if match_pza:
            return int(match_pza.group(1))
        if "CAJ" in u or "CAJA" in u:
            if "12" in combined: return 12
            if "6" in combined: return 6
            if "24" in combined: return 24
            if "48" in combined: return 48
        return 1

    # Regla general / PriceSmart
    match_pza = re.search(r'(\d+)\s*PZA', combined)
    if match_pza:
        return int(match_pza.group(1))

    match_slash = re.search(r'\b(\d+)\s*/', combined)
    if match_slash:
        val = int(match_slash.group(1))
        if val in [1, 3, 6, 12, 16, 20, 24, 48]:
            return val

    if "CAJ" in u or "CAJA" in u:
        if "12" in combined: return 12
        if "6" in combined: return 6
        if "24" in combined: return 24
        if "48" in combined: return 48

    return 1

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Memoria Inteligente de Proveedores</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📁 Actualizar Catálogo Maestro", "🏢 Perfiles de Proveedores", "📋 Códigos Almacenados"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Configuración de API")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

def process_invoice_smart_router(file_obj, file_type, use_paid_gemini=False):
    prompt_detect = (
        "Identifica el nombre comercial del proveedor en esta factura. "
        "Devuelve un JSON puro: {'proveedor': 'NOMBRE_PROVEEDOR'}"
    )

    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini:
        active_key = ACTIVE_GEMINI_FREE_KEY

    genai.configure(api_key=active_key if active_key else ACTIVE_GEMINI_FREE_KEY)
    model = genai.GenerativeModel('gemini-3.6-flash')
    
    file_obj.seek(0)
    file_bytes = file_obj.read()
    
    if "pdf" in file_type.lower():
        image_input = {"mime_type": "application/pdf", "data": file_bytes}
    else:
        image_input = Image.open(io.BytesIO(file_bytes))

    supplier_detected = "GENERAL"
    try:
        resp_det = model.generate_content([image_input, prompt_detect])
        txt_det = resp_det.text.strip()
        if txt_det.startswith("```json"): txt_det = txt_det[7:]
        if txt_det.endswith("```"): txt_det = txt_det[:-3]
        j_det = json.loads(txt_det.strip())
        supplier_detected = str(j_det.get("proveedor") or "GENERAL").upper().strip()
    except Exception:
        supplier_detected = "GENERAL"

    is_formal_supplier = any(k in supplier_detected for k in ["GONZALEZ", "ROYAL", "CUESTA"])
    
    if is_formal_supplier:
        prompt_main = (
            "Analiza esta factura formal de proveedor con doble código en la izquierda. "
            "REGLA: Extrae estrictamente el **código de barras EAN (el número largo de 13 dígitos)** que está debajo del SAP. "
            "Extrae un JSON bajo la clave 'items': "
            '{"proveedor": "' + supplier_detected + '", "items": [{"codigo_producto": "7804300150082", "descripcion": "...", "cantidad": 1.0, "umv": "CAJ / 12 PZA", "valor": 0.0}]}. '
            "Respuesta JSON pura sin texto adicional."
        )
    else:
        prompt_main = (
            "Analiza esta factura o tiquet de proveedor. "
            "REGLA: Copia el nombre literal exacto del producto y su código de ítem principal. "
            "Extrae un JSON bajo la clave 'items': "
            '{"proveedor": "' + supplier_detected + '", "items": [{"codigo_producto": "...", "descripcion": "...", "cantidad": 1.0, "umv": "...", "valor": 0.0}]}. '
            "Respuesta JSON pura sin texto adicional."
        )

    for intento in range(2):
        try:
            file_obj.seek(0)
            response = model.generate_content([image_input, prompt_main])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            
            parsed = json.loads(raw_text.strip())
            
            # Registrar automáticamente en memoria si el proveedor es nuevo
            supplier_memory = st.session_state["supplier_memory"]
            if supplier_detected not in supplier_memory and supplier_detected != "GENERAL":
                supplier_memory[supplier_detected] = {
                    "nombre": supplier_detected,
                    "formato_codigo": "adaptativo_auto_aprendido",
                    "regla_empaque": "Detección inteligente según UMV y empaques.",
                    "descripcion": "Proveedor registrado automáticamente por el sistema."
                }
                save_supplier_memory()

            return parsed, f"✅ Éxito ({supplier_detected})"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower():
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
            
    return None, last_err

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    try:
        st.markdown("<h2>📊 Automatizador con <span style='color: #0284c7;'>Memoria por Proveedor</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>El sistema detecta el proveedor y aplica su regla específica guardada.</p>", unsafe_allow_html=True)
        st.markdown("---")

        render_master_status_banner()

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        c_col1, _ = st.columns([1, 3])
        with c_col1:
            margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
        uploaded_file = st.file_uploader("📂 Sube tu factura o tiquet (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file is not None:
            st.success(f"¡Archivo cargado: {uploaded_file.name}!")
            if st.button("🚀 Procesar Factura"):
                file_type = getattr(uploaded_file, 'type', 'image/jpeg')
                with st.spinner("Procesando con perfil de proveedor detectado..."):
                    parsed_data, success_msg = process_invoice_smart_router(uploaded_file, file_type, use_paid_gemini=use_gemini_paid_api)

                if success_msg == "QUOTA_EXCEEDED":
                    st.error("⚠️ **Límite de cuota alcanzado.** Activa la casilla de 'Gemini Paid (API de Pago)' en la barra lateral.")
                elif not parsed_data or not isinstance(parsed_data, dict):
                    st.error(f"⚠️ Error al procesar el documento: {success_msg}")
                else:
                    prov_detectado = parsed_data.get("proveedor", "GENERAL")
                    st.success(f"{success_msg} | 🏢 Proveedor: **{prov_detectado}**")
                    
                    data_items = parsed_data.get("items", [])
                    rows_preview = []
                    multiplicador_ganancia = 1 + (margen_ganancia / 100.0)
                    calc_subtotal = 0.0

                    for idx, item in enumerate(data_items, start=1):
                        if not isinstance(item, dict):
                            continue
                        
                        desc_raw = str(item.get("descripcion") or "").strip()
                        if not desc_raw:
                            continue

                        invoice_code = str(item.get("codigo_producto") or "")
                        resolved_name, resolved_code = get_resolved_code_and_name(desc_raw, invoice_code)

                        cant_comprada = safe_float(item.get("cantidad") or 1, 1.0)
                        val_neto_linea = safe_float(item.get("valor") or 0)
                        umv_txt = str(item.get("umv") or "")
                        
                        empaque_val = parse_empaque_by_supplier(prov_detectado, umv_txt, resolved_name)
                        total_unidades = int(cant_comprada * empaque_val)

                        costo = round(val_neto_linea / total_unidades, 2) if total_unidades > 0 else 0.0
                        if costo <= 0:
                            continue

                        calc_subtotal += val_neto_linea
                        raw_pv = (costo * multiplicador_ganancia) * 1.18
                        precio_venta = round_to_nearest_5(raw_pv)
                        
                        rows_preview.append({
                            "No.": idx,
                            "Código Oficial / EAN": str(resolved_code),
                            "Nombre Factura / Artículo": resolved_name,
                            "Cant. Compra": cant_comprada,
                            "UMV": umv_txt,
                            "Empaque": empaque_val,
                            "Stock Total Unidades": total_unidades,
                            "Costo Unitario": costo,
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
                        st.markdown("### ✅ Artículos Procesados Exitosamente")
                        df_resultado = pd.DataFrame(rows_preview)
                        df_resultado["Código Oficial / EAN"] = df_resultado["Código Oficial / EAN"].astype(str)
                        
                        altura_tabla = min(max(len(rows_preview) * 35 + 40, 200), 850)
                        st.dataframe(df_resultado, use_container_width=True, hide_index=True, height=altura_tabla)
                        
                        wb = openpyxl.Workbook()
                        ws_prod = wb.active
                        ws_prod.title = "Productos"
                        ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                        
                        for item_dict in rows_preview:
                            ws_prod.append([
                                item_dict["Nombre Factura / Artículo"],
                                str(item_dict["Código Oficial / EAN"]),
                                "General", "producto",
                                item_dict["Precio Venta"],
                                item_dict["Costo Unitario"],
                                item_dict["Stock Total Unidades"],
                                5, 0.18, "unidad", "No",
                                item_dict["Empaque"], "No", 0, 0, None, "No", None
                            ])
                            ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                        
                        output = io.BytesIO()
                        wb.save(output)
                        st.download_button("📥 Descargar Excel WilPOS Oficial", output.getvalue(), f"Inventario_{prov_detectado.replace(' ', '_')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error("⚠️ Ocurrió un error inesperado en Factura Individual:")
        st.exception(e)

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    try:
        st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes y Consolidación con Memoria</span></h2>", unsafe_allow_html=True)
        st.markdown("---")
        render_master_status_banner()
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        l_col1, _ = st.columns([1, 3])
        with l_col1:
            margen_ganancia_lote = st.number_input("⚙️ Ganancia (%) Lote", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        uploaded_files = st.file_uploader("📂 Sube tus facturas (Selección múltiple)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_files:
            if "cached_uploaded_files" not in st.session_state or len(st.session_state["cached_uploaded_files"]) != len(uploaded_files):
                st.session_state["cached_uploaded_files"] = [
                    {"name": f.name, "type": getattr(f, "type", "image/jpeg"), "bytes": f.read()}
                    for f in uploaded_files
                ]

        if "cached_uploaded_files" in st.session_state and st.session_state["cached_uploaded_files"]:
            cached_files = st.session_state["cached_uploaded_files"]
            total_files = len(cached_files)

            if "batch_accumulated_items" not in st.session_state:
                st.session_state["batch_accumulated_items"] = []
                st.session_state["batch_audit_log"] = []
                st.session_state["batch_processed_count"] = 0
                st.session_state["batch_ok_count"] = 0
                st.session_state["is_live_processing"] = False
                st.session_state["quota_paused"] = False

            processed_so_far = st.session_state["batch_processed_count"]

            if st.session_state.get("quota_paused", False):
                st.error("⚠️ **Límite de cuota alcanzado.**")
                if st.button("💎 Reintentar con Gemini Paid", type="primary"):
                    st.session_state["quota_paused"] = False
                    st.session_state["is_live_processing"] = True
                    st.rerun()
            else:
                b_col1, b_col2 = st.columns(2)
                iniciar_btn = b_col1.button("🚀 Iniciar Lote y Consolidar", type="primary")
                reiniciar_lote = b_col2.button("🔄 Reiniciar Lote")

                if reiniciar_lote:
                    st.session_state["batch_accumulated_items"] = []
                    st.session_state["batch_audit_log"] = []
                    st.session_state["batch_processed_count"] = 0
                    st.session_state["batch_ok_count"] = 0
                    st.session_state["is_live_processing"] = False
                    st.session_state["quota_paused"] = False
                    if "cached_uploaded_files" in st.session_state:
                        del st.session_state["cached_uploaded_files"]
                    st.rerun()

                if iniciar_btn:
                    st.session_state["is_live_processing"] = True
                    st.rerun()

            is_live = st.session_state.get("is_live_processing", False)

            if is_live and not st.session_state.get("quota_paused", False):
                if processed_so_far < total_files:
                    file_info = cached_files[processed_so_far]
                    st.info(f"⚡ Procesando {processed_so_far + 1} de {total_files}: `{file_info['name']}`...")
                    st.progress(processed_so_far / total_files)

                    file_bytes_io = io.BytesIO(file_info["bytes"])
                    parsed_data, err_msg = process_invoice_smart_router(file_bytes_io, file_info["type"], use_paid_gemini=use_gemini_paid_api)
                    time.sleep(1.0)

                    if err_msg == "QUOTA_EXCEEDED":
                        st.session_state["is_live_processing"] = False
                        st.session_state["quota_paused"] = True
                        st.rerun()

                    if parsed_data and isinstance(parsed_data, dict):
                        st.session_state["batch_ok_count"] += 1
                        prov_det = parsed_data.get("proveedor", "GENERAL")
                        st.session_state["batch_audit_log"].append({"Archivo": file_info["name"], "Proveedor": prov_det, "Estado": "🟢 OK"})
                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            for itm in items:
                                if isinstance(itm, dict):
                                    d_txt = str(itm.get("descripcion") or "").strip()
                                    if d_txt:
                                        itm["_proveedor_detected"] = prov_det
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
                total_unidades_inventario = 0

                for item in raw_items:
                    if not isinstance(item, dict):
                        continue
                    
                    desc_raw = str(item.get("descripcion") or "").strip()
                    if not desc_raw:
                        continue

                    invoice_code = str(item.get("codigo_producto") or "")
                    prov_det = str(item.get("_proveedor_detected") or "GENERAL")
                    resolved_name, resolved_code = get_resolved_code_and_name(desc_raw, invoice_code)
                    
                    cant_comprada = safe_float(item.get("cantidad") or 1, 1.0)
                    val_neto_linea = safe_float(item.get("valor") or 0)
                    umv_txt = str(item.get("umv") or "")
                    
                    empaque_val = parse_empaque_by_supplier(prov_det, umv_txt, resolved_name)
                    total_unidades = int(cant_comprada * empaque_val)

                    costo = round(val_neto_linea / total_unidades, 2) if total_unidades > 0 else 0.0
                    if costo <= 0:
                        continue

                    raw_pv = (costo * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    total_unidades_inventario += total_unidades

                    processed_rows.append({
                        "Nombre": resolved_name,
                        "Código Barra": str(resolved_code),
                        "Categoría": "General", "Tipo": "producto",
                        "Precio Venta": precio_venta, "Costo": costo,
                        "Stock": total_unidades, "Stock Mínimo": 5, "ITBIS": 0.18,
                        "Unidad Medida": "unidad", "Venta Granel": "No",
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

                    df_final_preview = df_grouped.sort_values(by="Stock", ascending=False).reset_index(drop=True)
                    df_final_preview["Código Barra"] = df_final_preview["Código Barra"].astype(str)

                    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                    kpi1.metric("📁 Facturas Procesadas", f"{st.session_state['batch_ok_count']}")
                    kpi2.metric("📦 Total Líneas", f"{len(raw_items)}")
                    kpi3.metric("📦 Unidades en Stock", f"{total_unidades_inventario:,}")
                    
                    inversion_total_lote = (df_final_preview['Costo'] * df_final_preview['Stock']).sum()
                    kpi4.metric("💰 Inversión Neta", f"RD$ {inversion_total_lote:,.2f}")

                    st.markdown("---")
                    altura_tabla_lote = min(max(len(df_final_preview) * 35 + 40, 200), 850)
                    st.dataframe(df_final_preview, use_container_width=True, hide_index=True, height=altura_tabla_lote)

                    wb = openpyxl.Workbook()
                    ws_prod = wb.active
                    ws_prod.title = "Productos"
                    ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])

                    for _, row in df_final_preview.iterrows():
                        ws_prod.append([
                            row["Nombre"], str(row["Código Barra"]), row["Categoría"], row["Tipo"],
                            row["Precio Venta"], row["Costo"], row["Stock"], row["Stock Mínimo"],
                            row["ITBIS"], row["Unidad Medida"], row["Venta Granel"], row["Cantidad Empaque"],
                            row["Precio Variable"], row["Descuento %"], row["Descuento Monto"],
                            row["Precio Especial"], row["Descuento Activo"], row["Descuento Nota"]
                        ])
                        ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'

                    output = io.BytesIO()
                    wb.save(output)
                    st.download_button("📥 Descargar Excel Consolidado Final", output.getvalue(), "Inventario_WilPOS_Consolidado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error("⚠️ Ocurrió un error inesperado en Lotes:")
        st.exception(e)

# ==========================================
# MÓDULO 3: ACTUALIZAR CATÁLOGO MAESTRO
# ==========================================
elif modulo == "📁 Actualizar Catálogo Maestro":
    try:
        st.markdown("<h2>📁 Actualización y Gestión del <span style='color: #0284c7;'>Catálogo Maestro</span></h2>", unsafe_allow_html=True)
        render_master_status_banner()
        st.markdown("---")
        master_file = st.file_uploader("📂 Sube tu Catálogo Maestro actualizado (Excel .xlsx)", type=["xlsx"], key="master_upload")
        
        if master_file is not None:
            try:
                df_master = pd.read_excel(master_file)
                st.success("¡Archivo Excel leído con éxito!")
                st.write("Vista previa:", df_master.head())
                
                cols = df_master.columns.tolist()
                col_name = st.selectbox("Columna con el Nombre / Descripción del Producto", cols)
                col_code = st.selectbox("Columna con el Código Oficial", cols)
                
                if st.button("🔄 Sobrescribir y Actualizar Maestro en el Sistema"):
                    count = 0
                    temp_dict = {}
                    for _, row in df_master.iterrows():
                        p_name = str(row[col_name]).strip().upper()
                        p_code = clean_code(row[col_code])
                        if p_name and p_code != "S/C (Sin Código)":
                            temp_dict[p_name] = p_code
                            count += 1
                    
                    st.session_state["master_catalog"] = temp_dict
                    save_master_to_file()
                    
                    timestamp_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_meta_to_file(timestamp_actual, count)
                    
                    st.success(f"¡Catálogo maestro actualizado exitosamente con {count} productos!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error al procesar el archivo Excel: {e}")

        if st.session_state["master_catalog"]:
            st.markdown("---")
            col_act1, col_act2 = st.columns([3, 1])
            col_act1.markdown(f"### 📋 Productos activos en el Sistema ({len(st.session_state['master_catalog'])})")
            if col_act2.button("🗑️ Vaciar Catálogo"):
                st.session_state["master_catalog"] = {}
                save_master_to_file()
                save_meta_to_file("Nunca", 0)
                st.rerun()
                
            df_current_master = pd.DataFrame([{"Descripción": k, "Código Oficial": v} for k, v in st.session_state["master_catalog"].items()])
            st.dataframe(df_current_master, use_container_width=True, hide_index=True, height=400)
    except Exception as e:
        st.error("⚠️ Ocurrió un error inesperado en Catálogo Maestro:")
        st.exception(e)

# ==========================================
# MÓDULO 4: PERFILES DE PROVEEDORES
# ==========================================
elif modulo == "🏢 Perfiles de Proveedores":
    try:
        st.markdown("<h2>🏢 Memoria y Perfiles de <span style='color: #0284c7;'>Proveedores</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Visualiza los perfiles de formato registrados para cada proveedor en la memoria del sistema.</p>", unsafe_allow_html=True)
        st.markdown("---")

        suppliers = st.session_state["supplier_memory"]
        if suppliers:
            st.success(f"🟢 Hay **{len(suppliers)} proveedores** registrados en la memoria persistente.")
            for prov_name, prov_data in suppliers.items():
                with st.expander(f"🏢 {prov_name}"):
                    st.write(f"**Formato de Código:** {prov_data.get('formato_codigo', 'N/A')}")
                    st.write(f"**Regla de Empaque:** {prov_data.get('regla_empaque', 'N/A')}")
                    st.write(f"**Descripción del Proveedor:** {prov_data.get('descripcion', 'N/A')}")
        else:
            st.info("ℹ️ No hay proveedores registrados.")
    except Exception as e:
        st.error("⚠️ Ocurrió un error en Perfiles de Proveedores:")
        st.exception(e)

# ==========================================
# MÓDULO 5: CÓDIGOS ALMACENADOS
# ==========================================
elif modulo == "📋 Códigos Almacenados":
    try:
        st.markdown("<h2>📋 Memoria de <span style='color: #0284c7;'>Códigos y Catálogo</span></h2>", unsafe_allow_html=True)
        render_master_status_banner()
        st.markdown("---")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown(f"### 📚 Catálogo Maestro ({len(st.session_state['master_catalog'])})")
            if st.session_state["master_catalog"]:
                df_m_mem = pd.DataFrame([{"Producto": k, "Código Oficial": v} for k, v in st.session_state["master_catalog"].items()])
                st.dataframe(df_m_mem, use_container_width=True, hide_index=True, height=400)
            else:
                st.warning("⚠️ No hay Catálogo Maestro cargado.")
                
        with col_m2:
            b_mem = st.session_state["barcode_memory"]
            st.markdown(f"### ⚡ Memoria de Aprendizaje ({len(b_mem)})")
            if b_mem:
                df_codes = pd.DataFrame([{"Código Oficial": str(code), "Producto": name} for name, code in b_mem.items()])
                st.dataframe(df_codes, use_container_width=True, hide_index=True, height=400)
            else:
                st.warning("⚠️ La memoria de aprendizaje está vacía.")
    except Exception as e:
        st.error("⚠️ Ocurrió un error en Códigos Almacenados:")
        st.exception(e)
