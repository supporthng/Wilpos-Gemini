import io
import json
import os
import hashlib
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

# Configuración de Claves API
free_key_1 = st.secrets.get("GEMINI_API_KEY_1", os.environ.get("GEMINI_API_KEY_1", ""))
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", os.environ.get("GEMINI_API_KEY_2", ""))
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", os.environ.get("GEMINI_API_KEY_PAID", ""))

# Archivos persistentes
MEMORY_FILE = "proveedores_memoria.json"
OVERRIDES_FILE = "mapeo_productos_overrides.json"
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

if "provider_memory" not in st.session_state:
    st.session_state["provider_memory"] = load_json_file(MEMORY_FILE)

if "product_overrides" not in st.session_state:
    st.session_state["product_overrides"] = load_json_file(OVERRIDES_FILE)

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

# Funciones auxiliares
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
    return round(round(x / 5) * 5, 2)

def clean_barcode(code_val):
    if not code_val:
        return "S/C (Sin Código)"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'):
        s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c"]:
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
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📸 Extraer Código desde Imagen", "📋 Ver Códigos Almacenados"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🧠 Sistema de Memoria")
with st.sidebar.expander(f"📦 Códigos Registrados ({len(st.session_state['barcode_memory'])} ítems)"):
    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.write(f"Total en memoria viva: {len(b_mem)}")
        if st.button("🗑️ Limpiar Memoria"):
            st.session_state["barcode_memory"] = {}
            if os.path.exists(BARCODE_MEMORY_FILE):
                os.remove(BARCODE_MEMORY_FILE)
            st.success("¡Reseteado!")
            st.rerun()
    else:
        st.info("Sin códigos guardados.")

with st.sidebar.expander("🛠️ Correcciones Manuales"):
    overrides = st.session_state["product_overrides"]
    if overrides:
        st.write(f"Reglas: {len(overrides)}")
        if st.button("🗑️ Limpiar Reglas"):
            st.session_state["product_overrides"] = {}
            if os.path.exists(OVERRIDES_FILE):
                os.remove(OVERRIDES_FILE)
            st.success("¡Reseteado!")
            st.rerun()
    else:
        st.info("Sin reglas manuales.")

if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL",
        "CORONA CERO": "CERVEZA CORONA CERO"
    }

# ==========================================
# MOTOR DE INTELIGENCIA Y EMPAREJAMIENTO
# ==========================================
def validate_with_master(item_description, original_code):
    clean_desc_key = str(item_description).strip().upper()
    
    if "CORONA CERO" in clean_desc_key or "CERO 355" in clean_desc_key:
        return "750304423180", "Actualizado (Regla Maestra Inmediata Corona Cero)"

    clean_orig_code = clean_barcode(original_code)

    if clean_desc_key in st.session_state["product_overrides"]:
        return clean_barcode(st.session_state["product_overrides"][clean_desc_key]), "Actualizado (Regla Guardada)"

    b_mem = st.session_state["barcode_memory"]
    for b_code, b_name in b_mem.items():
        if b_name == clean_desc_key:
            return clean_barcode(b_code), "Actualizado (Memoria Viva Exacta)"

    if b_mem:
        name_to_code = {str(name).upper(): code for code, name in b_mem.items()}
        close_matches = difflib.get_close_matches(clean_desc_key, name_to_code.keys(), n=1, cutoff=0.65)
        if close_matches:
            matched_name = close_matches[0]
            code_found = name_to_code[matched_name]
            return clean_barcode(code_found), f"Actualizado (Memoria Viva por Similitud: '{matched_name}')"

    if clean_orig_code != "S/C (Sin Código)":
        b_mem[clean_orig_code] = clean_desc_key
        save_json_file(BARCODE_MEMORY_FILE, b_mem)
        return clean_orig_code, "✨ Nuevo Código Registrado Automáticamente en Memoria"

    return "S/C (Sin Código)", "⚠️ Sin Código Detectado"

def audit_and_correct_cost(costo_unit, cantidad, empaque):
    c = safe_float(costo_unit)
    cant = safe_int(cantidad, 1)
    emp = safe_int(empaque, 1)
    if emp == 1 and cant > 1:
        return c
    if c > 5000 and (cant * emp) > 1:
        corrected = c / (cant * emp)
        if corrected > 0:
            return corrected
    return c

def process_invoice_with_ai(file_obj, file_type):
    memory_context = ""
    known_mem = st.session_state["provider_memory"]
    if known_mem:
        memory_context = "MEMORIA HISTÓRICA DE FORMATOS DE PROVEEDORES:\n"
        for rnc, info in known_mem.items():
            memory_context += f"- Proveedor RNC {rnc} ({info.get('nombre', '')}): {info.get('nota_formato', 'Formato estándar de cajas con empaques fraccionados.')}\n"

    prompt_text = (
        f"{memory_context}\n"
        "Analiza esta factura detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'emisor_nombre', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae: 'codigo' (ASEGÚRATE DE NO OMITIR NINGÚN CERO A LA IZQUIERDA Y DEVUÉLVELO EXACTAMENTE COMO TEXTO), 'descripcion', 'cantidad', 'empaque', y 'costo_sin_itbis'. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    keys_to_try = []
    if paid_api_key:
        keys_to_try.append((paid_api_key, "Versión de Pago (Paid Tier)"))
    if free_key_1:
        keys_to_try.append((free_key_1, "Respaldo Gratuito #1"))
    if free_key_2:
        keys_to_try.append((free_key_2, "Respaldo Gratuito #2"))

    if not keys_to_try:
        return None, ""

    for api_k, label in keys_to_try:
        try:
            genai.configure(api_key=api_k)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            response = model.generate_content([
                {'mime_type': file_type, 'data': file_bytes},
                prompt_text
            ])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
            parsed_data = json.loads(raw_text.strip())
            return parsed_data, f"✅ Éxito con {label}"
        except Exception as e:
            if "429" in str(e) or "Quota exceeded" in str(e):
                continue
            else:
                break
    return None, ""

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura individual con preservación estricta de códigos con ceros a la izquierda.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, c_col2 = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0, key="textbox_individual")
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            with st.spinner("Analizando factura..."):
                parsed_data, success_msg = process_invoice_with_ai(uploaded_file, file_type)

            if parsed_data:
                st.success(success_msg)
                data_items = parsed_data.get("items", [])
                rows_preview = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = clean_barcode(item.get("codigo", ""))
                    final_code, status_match = validate_with_master(desc, orig_code)

                    raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)
                    raw_pv = (costo * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = cant_comprada * empaque_val
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Barra POS": str(final_code),
                        "Nombre": desc,
                        "Cant. Compra": cant_comprada,
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unit. Sin ITBIS": costo,
                        "Precio Venta": precio_venta,
                        "Estado Memoria": status_match
                    })

                df_resultado = pd.DataFrame(rows_preview)
                df_resultado["Código Barra POS"] = df_resultado["Código Barra POS"].astype(str)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                wb = openpyxl.Workbook()
                ws_prod = wb.active
                ws_prod.title = "Productos"
                ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                
                for item_dict in rows_preview:
                    row_cells = [
                        item_dict["Nombre"],
                        str(item_dict["Código Barra POS"]),
                        "General",
                        "producto",
                        item_dict["Precio Venta"],
                        item_dict["Costo Unit. Sin ITBIS"],
                        item_dict["Stock Total"],
                        5,
                        0.18,
                        "unidad",
                        "No",
                        item_dict["Empaque"],
                        "No",
                        0,
                        0,
                        None,
                        "No",
                        None
                    ]
                    ws_prod.append(row_cells)
                    ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                
                output = io.BytesIO()
                wb.save(output)
                st.download_button("📥 Descargar Excel Plantilla WilPOS Actualizada", output.getvalue(), "Inventario_WilPOS_Actualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes con Dashboard de Auditoría</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas con control detallado de procesados, omitidos y motivos.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    l_col1, l_col2 = st.columns([1, 3])
    with l_col1:
        margen_ganancia_lote = st.number_input("⚙️ Ganancia (%) Lote", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    uploaded_files = st.file_uploader("📂 Sube tu factura (Selección múltiple)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files:
        st.info(f"📁 Se han cargado **{len(uploaded_files)} archivo(s)** listos para procesar.")
        if st.button("🚀 Procesar Lote y Mostrar Dashboard", type="primary"):
            all_consolidated_items = []
            archivos_procesados_ok = 0
            audit_log = []
            batch_signatures = set()
            
            progress_bar = st.progress(0)
            status_placeholder = st.empty()
            total_files = len(uploaded_files)

            for i, file in enumerate(uploaded_files):
                status_placeholder.markdown(f"⏳ **Procesando archivo ({i+1} de {total_files}):** `{file.name}`...")
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                parsed_data, err_msg = process_invoice_with_ai(file, file_type)

                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    nombre_prov = str(parsed_data.get("emisor_nombre", "Desconocido")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc_val = safe_float(parsed_data.get("total", 0))
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc_val}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in batch_signatures:
                        audit_log.append({
                            "Archivo": file.name,
                            "Proveedor": nombre_prov,
                            "Nº Documento": num_doc if num_doc else "N/D",
                            "Estado": "🔴 Omitido (Duplicado)",
                            "Motivo": f"Factura ya existente en el lote actual (RNC: {rnc_emisor}, Doc: {num_doc}, Total: RD$ {total_doc_val:,.2f})"
                        })
                    else:
                        batch_signatures.add(doc_signature)
                        archivos_procesados_ok += 1
                        audit_log.append({
                            "Archivo": file.name,
                            "Proveedor": nombre_prov,
                            "Nº Documento": num_doc if num_doc else "N/D",
                            "Estado": "🟢 Procesado Exitosamente",
                            "Motivo": f"Extraídos {len(parsed_data.get('items', []))} ítems correctamente."
                        })
                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            all_consolidated_items.extend(items)
                else:
                    audit_log.append({
                        "Archivo": file.name,
                        "Proveedor": "Desconocido",
                        "Nº Documento": "N/D",
                        "Estado": "🔴 Omitido (Error de Lectura/IA)",
                        "Motivo": "La IA no pudo estructurar correctamente el documento o se agotó la cuota de la API."
                    })

                progress_bar.progress((i + 1) / total_files)

            status_placeholder.success("🎉 ¡Procesamiento y auditoría de lote finalizados!")

            st.markdown("---")
            st.markdown("## 📊 Dashboard de Auditoría y Procesamiento")
            
            total_subidos = len(uploaded_files)
            total_omitidos = total_subidos - archivos_procesados_ok

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("📁 Archivos Subidos", total_subidos)
            col_m2.metric("🟢 Procesados OK", archivos_procesados_ok)
            col_m3.metric("🔴 Omitidos / Filtrados", total_omitidos)
            col_m4.metric("📦 Productos Totales", len(all_consolidated_items))

            st.markdown("### 📋 Detalle de Archivos Evaluados (Procesados vs Omitidos)")
            df_audit = pd.DataFrame(audit_log)
            st.dataframe(df_audit, use_container_width=True, hide_index=True)

            if all_consolidated_items:
                st.markdown("---")
                st.markdown(f"### 📦 Consolidado de Inventario Resultante")

                rows_preview = []
                multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)

                for idx, item in enumerate(all_consolidated_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = clean_barcode(item.get("codigo", ""))
                    final_code, status_match = validate_with_master(desc, orig_code)

                    raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)
                    raw_pv = (costo * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = cant_comprada * empaque_val
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Barra POS": str(final_code),
                        "Nombre": desc,
                        "Cant. Compra": cant_comprada,
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unit. Sin ITBIS": costo,
                        "Precio Venta": precio_venta,
                        "Estado Memoria": status_match
                    })

                df_batch = pd.DataFrame(rows_preview)
                df_batch["Código Barra POS"] = df_batch["Código Barra POS"].astype(str)
                st.dataframe(df_batch, use_container_width=True, hide_index=True)

                wb = openpyxl.Workbook()
                ws_prod = wb.active
                ws_prod.title = "Productos"
                ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])

                for item_dict in rows_preview:
                    ws_prod.append([
                        item_dict["Nombre"],
                        str(item_dict["Código Barra POS"]),
                        "General",
                        "producto",
                        item_dict["Precio Venta"],
                        item_dict["Costo Unit. Sin ITBIS"],
                        item_dict["Stock Total"],
                        5,
                        0.18,
                        "unidad",
                        "No",
                        item_dict["Empaque"],
                        "No",
                        0,
                        0,
                        None,
                        "No",
                        None
                    ])
                    ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'

                output = io.BytesIO()
                wb.save(output)
                st.download_button("📥 Descargar Excel Consolidado Sin Duplicados", output.getvalue(), "Inventario_WilPOS_Consolidado_Lote.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# MÓDULO 3: EXTRAER CÓDIGO DESDE IMAGEN
# ==========================================
elif modulo == "📸 Extraer Código desde Imagen":
    st.markdown("<h2>📸 Lector de Códigos y <span style='color: #0284c7;'>Productos</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube una foto del producto. Los ceros a la izquierda se respetarán de forma exacta.</p>", unsafe_allow_html=True)
    st.markdown("---")
    img_uploaded = st.file_uploader("📂 Sube la imagen del producto", type=["png", "jpg", "jpeg", "webp"], key="barcode_img_upload")
    if img_uploaded is not None:
        st.image(img_uploaded, caption="Imagen analizada", width=400)
        if st.button("🔍 Escanear y Registrar"):
            st.info("Escáner rápido disponible en la versión completa.")

# ==========================================
# MÓDULO 4: VER CÓDIGOS ALMACENADOS E IMPORTAR EXCEL
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.markdown("<h2>📋 Memoria Viva de <span style='color: #0284c7;'>Códigos Almacenados</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Gestión de códigos preservando cada cero inicial como texto plano.</p>", unsafe_allow_html=True)
    st.markdown("---")

    tab_view, tab_import_excel = st.tabs(["📊 Ver Almacenados", "📂 Importar/Actualizar desde Excel"])

    with tab_view:
        b_mem = st.session_state["barcode_memory"]
        if b_mem:
            st.info(f"📊 Total de códigos aprendidos y almacenados: **{len(b_mem)}**")
            df_codes = pd.DataFrame([{"Código de Barras": str(code), "Nombre del Producto": name} for code, name in b_mem.items()])
            df_codes["Código de Barras"] = df_codes["Código de Barras"].astype(str)
            st.dataframe(df_codes, use_container_width=True, hide_index=True, height=450)

            output_db = io.BytesIO()
            with pd.ExcelWriter(output_db, engine='openpyxl') as writer:
                df_codes.to_excel(writer, index=False, sheet_name="Codigos_Almacenados")
            
            wb_db = openpyxl.load_workbook(output_db)
            ws_db = wb_db.active
            for r in range(2, ws_db.max_row + 1):
                ws_db.cell(row=r, column=1).number_format = '@'
            final_db_output = io.BytesIO()
            wb_db.save(final_db_output)

            st.download_button("📥 Descargar Tu Memoria Actualizada (.xlsx)", final_db_output.getvalue(), "codigos_barras_almacenados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ La memoria está vacía.")

    with tab_import_excel:
        st.subheader("📂 Importación Masiva desde Excel Maestro")
        st.markdown("Al importar, los códigos de barras se leen explícitamente como texto para conservar todos los ceros a la izquierda.")
        
        # 🔄 BOTÓN DE REINICIO RÁPIDO PARA PRUEBAS
        if st.button("🔄 Reiniciar Memoria (Borrar Todo) y Subir Nuevo Excel"):
            st.session_state["barcode_memory"] = {}
            if os.path.exists(BARCODE_MEMORY_FILE):
                os.remove(BARCODE_MEMORY_FILE)
            if "last_uploaded_excel_name" in st.session_state:
                del st.session_state["last_uploaded_excel_name"]
            if "df_imported" in st.session_state:
                del st.session_state["df_imported"]
            st.success("¡Memoria reseteada con éxito! Ya puedes subir un nuevo archivo Excel.")
            st.rerun()

        st.markdown("---")
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        excel_import_file = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx", "xls", "csv"], key="import_memory_file")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if excel_import_file is not None:
            try:
                if "last_uploaded_excel_name" not in st.session_state or st.session_state["last_uploaded_excel_name"] != excel_import_file.name:
                    st.session_state["last_uploaded_excel_name"] = excel_import_file.name
                    if excel_import_file.name.endswith('.csv'):
                        st.session_state["df_imported"] = pd.read_csv(excel_import_file, dtype=str)
                    else:
                        st.session_state["df_imported"] = pd.read_excel(excel_import_file, sheet_name=0, dtype=str)

                df_imp = st.session_state["df_imported"]
                
                st.markdown(f"**Vista previa del archivo cargado ({len(df_imp)} filas totales):**")
                st.dataframe(df_imp.head(10), use_container_width=True, hide_index=True)
                
                cols_lower = [str(c).lower() for c in df_imp.columns]
                c_code = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'codigo' in c or 'barra' in c or 'barcode' in c), None)
                c_name = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'nombre' in c or 'descripcion' in c), None)
                
                if c_code and c_name:
                    if st.button("📥 Sincronizar y Guardar en Memoria", type="primary"):
                        nuevos = 0
                        actualizados = 0
                        omitidos = 0
                        log_omitidos = []

                        for idx, row in df_imp.iterrows():
                            c_val = clean_barcode(row[c_code])
                            n_val = str(row[c_name]).strip().upper()
                            
                            if c_val == "S/C (Sin Código)" or not c_val:
                                omitidos += 1
                                log_omitidos.append({"Fila": idx + 2, "Nombre": n_val if n_val else "N/A", "Motivo": "Código de barras ausente o inválido"})
                                continue
                            if not n_val or n_val in ["NAN", "NONE", ""]:
                                omitidos += 1
                                log_omitidos.append({"Fila": idx + 2, "Código Barra": c_val, "Motivo": "Nombre de producto vacío o inválido"})
                                continue
                                
                            if c_val in st.session_state["barcode_memory"]:
                                if st.session_state["barcode_memory"][c_val] != n_val:
                                    st.session_state["barcode_memory"][c_val] = n_val
                                    actualizados += 1
                            else:
                                st.session_state["barcode_memory"][c_val] = n_val
                                nuevos += 1
                                
                        save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                        
                        st.success(f"🎯 ¡Sincronización completada con éxito protegiendo los ceros a la izquierda!")
                        
                        col_r1, col_r2, col_r3 = st.columns(3)
                        col_r1.metric("✨ Nuevos Agregados", nuevos)
                        col_r2.metric("🔄 Actualizados", actualizados)
                        col_r3.metric("⚠️ Omitidos / Descartados", omitidos)

                        if log_omitidos:
                            st.markdown("### 📋 Detalle de Productos Omitidos y Motivo")
                            df_omit = pd.DataFrame(log_omitidos)
                            st.dataframe(df_omit, use_container_width=True, hide_index=True)
                        else:
                            st.info("✅ Excelente: No hubo ningún producto omitido; todas las filas del Excel fueron procesadas con éxito.")
                else:
                    st.error("❌ No se pudieron detectar automáticamente las columnas 'Código Barra' y 'Nombre' en el archivo.")
            except Exception as ex:
                st.error(f"Error procesando el archivo: {ex}")
