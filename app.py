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

# Configuración de la página
st.set_page_config(page_title="WilPOS - Automatizador de Facturas", page_icon="📊", layout="wide")

# Configuración de Claves API desde secrets de Streamlit o variables de entorno
free_key_1 = st.secrets.get("GEMINI_API_KEY_1", os.environ.get("GEMINI_API_KEY_1", ""))
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", os.environ.get("GEMINI_API_KEY_2", ""))
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", os.environ.get("GEMINI_API_KEY_PAID", ""))

# Archivos persistentes de memoria y reglas
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

# Funciones auxiliares de cálculo y formato
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

if "quota_exceeded" not in st.session_state:
    st.session_state["quota_exceeded"] = False

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.title("Menú de Navegación")
modulo = st.sidebar.radio(
    "Selecciona el Módulo",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📸 Extraer Código desde Imagen", "📋 Ver Códigos Almacenados"]
)

st.sidebar.markdown("---")
st.sidebar.title("⚙️ Configuración de Precios")
margen_ganancia = st.sidebar.slider(
    "Porcentaje de Ganancia (%)", 
    min_value=0.0, 
    max_value=100.0, 
    value=25.0, 
    step=1.0, 
    help="Margen de ganancia aplicado sobre el costo para calcular el precio de venta antes de impuestos."
)

st.sidebar.markdown("---")
st.sidebar.title("🧠 Memoria y Reglas POS")
with st.sidebar.expander("Ver Códigos Escaneados Guardados"):
    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.write(f"Total códigos en memoria: {len(b_mem)}")
        for b_code, b_name in b_mem.items():
            st.markdown(f"- `{b_code}` ➔ **{b_name}**")
        if st.button("🗑️ Limpiar Memoria de Códigos"):
            st.session_state["barcode_memory"] = {}
            if os.path.exists(BARCODE_MEMORY_FILE):
                os.remove(BARCODE_MEMORY_FILE)
            st.success("¡Memoria de códigos reseteada!")
            st.rerun()
    else:
        st.info("No hay códigos guardados en memoria aún.")

with st.sidebar.expander("Ver Correcciones de Productos"):
    overrides = st.session_state["product_overrides"]
    if overrides:
        st.write(f"Total reglas de mapeo: {len(overrides)}")
        for prov_desc, pos_code in overrides.items():
            st.markdown(f"- `{prov_desc}` ➔ Código: **{pos_code}**")
        if st.button("🗑️ Limpiar Reglas de Mapeo"):
            st.session_state["product_overrides"] = {}
            if os.path.exists(OVERRIDES_FILE):
                os.remove(OVERRIDES_FILE)
            st.success("¡Reglas reseteadas!")
            st.rerun()
    else:
        st.info("No hay reglas manuales registradas.")

# Equivalencias personalizables
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL",
        "CORONA CERO": "CERVEZA CORONA CERO"
    }

st.sidebar.markdown("---")
st.sidebar.title("🔄 Equivalencias y Sinónimos")
with st.sidebar.expander("Ver / Editar Equivalencias"):
    eq_key = st.text_input("Término del Proveedor")
    eq_val = st.text_input("Equivalente en tu POS")
    if st.button("➕ Agregar Regla"):
        if eq_key and eq_val:
            st.session_state["custom_equivalences"][eq_key.strip().upper()] = eq_val.strip().upper()
            st.success("¡Regla agregada!")
            st.rerun()
            
    if st.session_state["custom_equivalences"]:
        st.markdown("**Reglas activas:**")
        to_remove = []
        for k, v in st.session_state["custom_equivalences"].items():
            if st.checkbox(f"{k} ➔ {v}", value=True, key=f"eq_{k}") == False:
                to_remove.append(k)
        if to_remove:
            for r in to_remove:
                del st.session_state["custom_equivalences"][r]
            st.rerun()

# ==========================================
# MOTOR DE INTELIGENCIA Y EMPAREJAMIENTO
# ==========================================
SYNONYMS_MAP = {
    "JW ": "JOHNNIE WALKER ",
    "JW.": "JOHNNIE WALKER",
    "BUCH ": "BUCHANANS ",
    "BUCHANAN": "BUCHANANS",
    "0.5 LT": "500ML",
    "0.5LT": "500ML",
    "1 LT": "1000ML",
    "1LT": "1000ML",
    "RON ": "",
    "TEQ ": "TEQUILA ",
    "BOTELLA ": "",
    "BOTELL ": ""
}

def clean_and_normalize(text):
    upper_text = str(text).upper()
    for prov, pos in st.session_state["custom_equivalences"].items():
        if prov in upper_text:
            upper_text = upper_text.replace(prov, pos)
            
    for abbr, full in SYNONYMS_MAP.items():
        upper_text = upper_text.replace(abbr, full)
        
    cleaned = re.sub(r'[^A-Z0-9\s]', ' ', upper_text)
    stopwords = {"DE", "EL", "LA", "LOS", "LAS", "Y", "EN", "UN", "UNA", "CON", "CL", "ML", "L", "BCA", "BOT"}
    tokens = [t for t in cleaned.split() if t not in stopwords]
    return tokens, upper_text

def validate_with_master(item_description, original_code):
    clean_desc_key = str(item_description).strip().upper()
    
    if "CORONA CERO" in clean_desc_key or "CERO 355" in clean_desc_key:
        return "750304423180", "Actualizado (Regla Maestra Inmediata Corona Cero)"

    clean_orig_code = str(original_code).strip()
    if clean_orig_code.endswith('.0'):
        clean_orig_code = clean_orig_code[:-2]
    if not clean_orig_code or clean_orig_code.lower() in ["nan", "none", ""]:
        clean_orig_code = ""

    if clean_desc_key in st.session_state["product_overrides"]:
        return str(st.session_state["product_overrides"][clean_desc_key]).strip(), "Actualizado (Regla Guardada)"

    for b_code, b_name in st.session_state["barcode_memory"].items():
        if b_name == clean_desc_key or clean_desc_key in b_name or b_name in clean_desc_key:
            return str(b_code).strip(), "Actualizado (Memoria de Códigos)"

    if not clean_orig_code:
        return "S/C (Sin Código)", "⚠️ Sin Código Original en Factura"

    return clean_orig_code, "⚠️ Conserva Código Original"

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
        "Analiza esta factura detalladamente. Extrae los datos de cabecera con absoluta precisión: 'emisor_rnc', 'emisor_nombre', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad', 'empaque', y 'costo_sin_itbis'. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    parsed_data = None
    success_msg = ""
    
    keys_to_try = []
    if paid_api_key:
        keys_to_try.append((paid_api_key, "Versión de Pago (Paid Tier)"))
    if free_key_1:
        keys_to_try.append((free_key_1, "Respaldo Gratuito #1"))
    if free_key_2:
        keys_to_try.append((free_key_2, "Respaldo Gratuito #2"))

    if not keys_to_try:
        st.error("❌ No se encontró ninguna clave de API configurada en los secrets de Streamlit.")
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
            success_msg = f"✅ ¡Factura procesada con éxito usando {label}!"
            return parsed_data, success_msg
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "Quota exceeded" in err_str:
                continue
            else:
                st.error(f"Error al procesar con {label}: {e}")
                break

    st.error("🚨 Se ha agotado la cuota de todas las claves configuradas.")
    return None, ""

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura para extraer sus ítems, validar códigos con tu memoria POS y generar la plantilla actualizada.")

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")

        with st.expander("👁️ Vista Previa del Archivo Cargado"):
            file_type_check = uploaded_file.type if hasattr(uploaded_file, 'type') else ''
            if "image" in file_type_check or uploaded_file.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                image = Image.open(uploaded_file)
                st.image(image, caption=f"Vista previa: {uploaded_file.name}", use_container_width=True)
                uploaded_file.seek(0)
            else:
                st.info(f"El archivo '{uploaded_file.name}' es de tipo PDF o documento.")

        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            
            with st.spinner("Analizando factura con auditoría de costos..."):
                parsed_data, success_msg = process_invoice_with_ai(uploaded_file, file_type)

            if parsed_data:
                st.success(success_msg)
                
                prov_nombre = parsed_data.get("emisor_nombre", "Desconocido")
                prov_rnc = parsed_data.get("emisor_rnc", "N/D")
                st.info(f"🏢 **Proveedor Procesado:** {prov_nombre} | **RNC:** `{prov_rnc}`")

                st.markdown("### 📋 Resumen de Totales de la Factura")
                c_t1, c_t2, c_t3 = st.columns(3)
                c_t1.metric("Subtotal", f"RD$ {safe_float(parsed_data.get('subtotal', 0)):,.2f}")
                c_t2.metric("ITBIS", f"RD$ {safe_float(parsed_data.get('itbis', 0)):,.2f}")
                c_t3.metric("Total General", f"RD$ {safe_float(parsed_data.get('total', 0)):,.2f}")
                
                st.markdown("---")
                st.markdown(f"### 📦 Validación con Memoria y Precios de Venta (Margen de Ganancia: {margen_ganancia}%)")

                data_items = parsed_data.get("items", [])
                rows_preview = []
                unmatched_items = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match or "Conserva" in status_match or "Sin Código" in status_match:
                        unmatched_items.append((idx, desc, orig_code))

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
                
                if unmatched_items:
                    st.warning(f"⚠️ **Atención:** Hay {len(unmatched_items)} producto(s) sin match automático:")
                    for u_idx, u_desc, u_code in unmatched_items:
                        st.markdown(f"- *{u_desc}*")
                        with st.expander(f"➕ Asignar Código POS correcto para: {u_desc} (Ítem #{u_idx})"):
                            new_pos_code = st.text_input(f"Introduce el código POS correcto", key=f"override_{u_idx}_{u_code}")
                            if st.button("Guardar Regla Mapeo", key=f"btn_override_{u_idx}_{u_code}"):
                                if new_pos_code:
                                    st.session_state["product_overrides"][u_desc.strip().upper()] = new_pos_code.strip()
                                    save_json_file(OVERRIDES_FILE, st.session_state["product_overrides"])
                                    st.success("¡Regla guardada con éxito! Vuelve a procesar para aplicarla.")
                                    st.rerun()

                df_resultado = pd.DataFrame(rows_preview)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
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
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Descargar Excel Plantilla WilPOS Actualizada",
                    data=excel_data,
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.title("📂 Procesador por Lotes")
    st.markdown(f"Sube varias facturas. El sistema validará los ítems con tu memoria de códigos aplicando un margen de ganancia del **{margen_ganancia}%**.")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        if st.button("🚀 Procesar Lote y Validar"):
            all_consolidated_items = []
            invoice_totals_summary = []
            duplicate_count = 0
            batch_signatures = set()

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}...")
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                
                parsed_data, _ = process_invoice_with_ai(file, file_type)

                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    nombre_emisor = str(parsed_data.get("emisor_nombre", "Desconocido")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    subtotal_doc = safe_float(parsed_data.get("subtotal", 0))
                    itbis_doc = safe_float(parsed_data.get("itbis", 0))
                    total_doc_val = safe_float(parsed_data.get("total", 0))
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc_val}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in batch_signatures:
                        duplicate_count += 1
                    else:
                        batch_signatures.add(doc_signature)
                        invoice_totals_summary.append({
                            "Proveedor": nombre_emisor if nombre_emisor and nombre_emisor != "None" else f"RNC: {rnc_emisor}",
                            "Archivo": file.name,
                            "Nº Documento": num_doc if num_doc else "N/D",
                            "Subtotal": subtotal_doc,
                            "ITBIS": itbis_doc,
                            "Total General": total_doc_val
                        })

                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            all_consolidated_items.extend(items)

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.text("¡Procesamiento por lotes completado!")

            if all_consolidated_items:
                st.markdown("---")
                st.markdown(f"### 📦 Consolidado de Ítems ({len(all_consolidated_items)} productos totales)")

                rows_preview = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                for idx, item in enumerate(all_consolidated_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
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
                excel_data_batch = output.getvalue()

                st.download_button(
                    label="📥 Descargar Excel Consolidado Sin Duplicados",
                    data=excel_data_batch,
                    file_name="Inventario_WilPOS_Consolidado_Lote.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MÓDULO 3: EXTRAER CÓDIGO DESDE IMAGEN
# ==========================================
elif modulo == "📸 Extraer Código desde Imagen":
    st.title("📸 Lector de Códigos y Productos (Memoria e Internet)")
    st.markdown("Sube la foto del código de barras o producto. Si no está en tu memoria, el sistema lo buscará en internet y lo registrará.")

    img_uploaded = st.file_uploader("Sube la imagen del producto (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg", "webp"], key="barcode_img_upload")

    if img_uploaded is not None:
        st.success(f"Imagen cargada: {img_uploaded.name}")
        
        col_prev1, col_prev2 = st.columns([1, 1])
        with col_prev1:
            image_obj = Image.open(img_uploaded)
            st.image(image_obj, caption="Imagen analizada", use_container_width=True)
            img_uploaded.seek(0)
            
        with col_prev2:
            if st.button("🔍 Escanear y Buscar (Memoria e Internet)"):
                keys_to_try = []
                if paid_api_key:
                    keys_to_try.append((paid_api_key, "Versión de Pago"))
                if free_key_1:
                    keys_to_try.append((free_key_1, "Respaldo #1"))
                if free_key_2:
                    keys_to_try.append((free_key_2, "Respaldo #2"))
                
                if not keys_to_try:
                    st.error("❌ No hay claves de API configuradas.")
                else:
                    extracted_code = ""
                    success = False
                    
                    prompt_scan = (
                        "Analiza esta imagen con código de barras o producto. "
                        "Extrae con absoluta precisión únicamente el número de código de barras visible. "
                        "Devuelve la respuesta estrictamente en formato JSON: "
                        '{"codigo_barras": "..."}'
                    )
                    
                    with st.spinner("Escaneando código de barras..."):
                        for api_k, label in keys_to_try:
                            try:
                                genai.configure(api_key=api_k)
                                model = genai.GenerativeModel('gemini-3.6-flash')
                                
                                img_uploaded.seek(0)
                                img_bytes = img_uploaded.read()
                                
                                resp = model.generate_content([
                                    {'mime_type': img_uploaded.type, 'data': img_bytes},
                                    prompt_scan
                                ])
                                
                                raw_t = resp.text.strip()
                                if raw_t.startswith("```json"):
                                    raw_t = raw_t[7:]
                                if raw_t.endswith("```"):
                                    raw_t = raw_t[:-3]
                                    
                                data_res = json.loads(raw_t.strip())
                                extracted_code = str(data_res.get("codigo_barras", "")).strip()
                                success = True
                                break
                            except Exception as e:
                                continue
                    
                    if success and extracted_code:
                        st.markdown("### 🎯 Resultado del Escaneo")
                        st.info(f"🔢 **Código de Barras Detectado:** `{extracted_code}`")
                        
                        if extracted_code in st.session_state["barcode_memory"]:
                            mem_name = st.session_state["barcode_memory"][extracted_code]
                            st.success("💾 ¡Encontrado en la memoria de códigos registrados!")
                            st.markdown(f"🏷️ **Nombre Registrado:** **{mem_name}**")
                        else:
                            st.warning("⚠️ Código nuevo. Consultando coincidencias en Internet...")
                            internet_product_name = "Producto Nuevo Escaneado"
                            with st.spinner("Buscando información del producto en internet..."):
                                try:
                                    search_model = genai.GenerativeModel('gemini-3.6-flash')
                                    search_resp = search_model.generate_content(
                                        f"Identifica y da el nombre comercial exacto, marca y presentación en español del producto cuyo código de barras universal es: {extracted_code}. Responde únicamente con el nombre del producto."
                                    )
                                    if search_resp and search_resp.text:
                                        internet_product_name = search_resp.text.strip().upper()
                                except Exception as ex:
                                    print(f"Error en búsqueda web: {ex}")
                            
                            st.session_state["barcode_memory"][extracted_code] = internet_product_name
                            save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                            
                            st.success("✨ ¡Producto nuevo identificado mediante internet y registrado en tu memoria!")
                            st.markdown(f"🏷️ **Nombre Identificado:** **{internet_product_name}**")
                    else:
                        st.error("No se pudo extraer un código de barras claro de la imagen.")

# ==========================================
# MÓDULO 4: VER CÓDIGOS ALMACENADOS
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.title("📋 Listado de Códigos y Nombres Almacenados")
    st.markdown("Consulta y alimenta tu memoria de códigos escaneados mediante carga masiva de Excel o lectura de imágenes.")

    tab_view, tab_import_excel, tab_import_image = st.tabs(["📊 Ver Almacenados", "📂 Extraer desde Excel", "📸 Leer desde Imagen"])

    with tab_view:
        b_mem = st.session_state["barcode_memory"]
        if b_mem:
            st.info(f"📊 Total de códigos registrados en memoria: **{len(b_mem)}**")
            list_data = [{"Código de Barras": code, "Nombre del Producto": name} for code, name in b_mem.items()]
            df_codes = pd.DataFrame(list_data)
            
            st.dataframe(df_codes, use_container_width=True, hide_index=True, height=450)

            csv_data = df_codes.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar Lista de Códigos (CSV)",
                data=csv_data,
                file_name="codigos_barras_almacenados.csv",
                mime="text/csv"
            )
        else:
            st.warning("⚠️ Aún no hay códigos de barras almacenados en la memoria.")

    with tab_import_excel:
        st.subheader("📂 Importar y Actualizar Códigos desde Excel o CSV")
        st.markdown("Sube tu archivo. Los códigos nuevos se agregarán y los existentes actualizarán su nombre automáticamente.")
        
        excel_import_file = st.file_uploader("Sube tu archivo Excel o CSV", type=["xlsx", "xls", "csv"], key="import_memory_file")
        
        if excel_import_file is not None:
            try:
                if excel_import_file.name.endswith('.csv'):
                    df_imp = pd.read_csv(excel_import_file, dtype=str)
                else:
                    df_imp = pd.read_excel(excel_import_file, dtype=str)
                
                st.markdown(f"**Vista previa completa del archivo cargado ({len(df_imp)} filas en total):**")
                st.dataframe(df_imp, use_container_width=True, hide_index=True, height=300)
                
                cols_lower = [str(c).lower() for c in df_imp.columns]
                c_code = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'codigo' in c or 'barra' in c or 'barcode' in c), None)
                c_name = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'nombre' in c or 'descripcion' in c), None)
                
                if c_code and c_name:
                    if st.button("📥 Importar, Actualizar y Agregar Nuevos a Memoria"):
                        total_filas = len(df_imp)
                        procesados_ok = 0
                        nuevos_agregados = 0
                        actualizados = 0
                        no_procesados = 0
                        motivos_no_procesados = []

                        for row_idx, row in df_imp.iterrows():
                            fila_num = row_idx + 2
                            c_val = str(row[c_code]).strip()
                            n_val = str(row[c_name]).strip().upper()
                            
                            nombre_articulo = n_val if (n_val and n_val.lower() not in ["nan", "none", ""]) else "(Sin Nombre / Artículo Desconocido)"
                            
                            if not c_val or c_val.lower() in ["nan", "none", ""]:
                                no_procesados += 1
                                motivos_no_procesados.append(f"Fila #{fila_num} ➔ **Artículo:** *{nombre_articulo}* | **Motivo:** Código de barras vacío o nulo.")
                                continue
                                
                            if not n_val or n_val.lower() in ["nan", "none", ""]:
                                no_procesados += 1
                                motivos_no_procesados.append(f"Fila #{fila_num} ➔ **Código:** `{c_val}` | **Motivo:** Nombre o descripción vacía.")
                                continue
                            
                            if c_val.endswith('.0'):
                                c_val = c_val[:-2]
                            
                            procesados_ok += 1
                            if c_val in st.session_state["barcode_memory"]:
                                if st.session_state["barcode_memory"][c_val] != n_val:
                                    st.session_state["barcode_memory"][c_val] = n_val
                                    actualizados += 1
                            else:
                                st.session_state["barcode_memory"][c_val] = n_val
                                nuevos_agregados += 1
                        
                        save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                        
                        st.success("🎯 **¡Proceso de importación finalizado!**")
                        col_r1, col_r2, col_r3, col_r4, col_r5 = st.columns(5)
                        col_r1.metric("Total Filas", total_filas)
                        col_r2.metric("Procesados", procesados_ok)
                        col_r3.metric("Nuevos Agregados", nuevos_agregados)
                        col_r4.metric("Actualizados", actualizados)
                        col_r5.metric("No Procesados", no_procesados)

                        if motivos_no_procesados:
                            with st.expander(f"⚠️ Ver detalle de los {len(motivos_no_procesados)} elementos no procesados y su artículo"):
                                for motivo in motivos_no_procesados:
                                    st.markdown(f"- {motivo}")
                else:
                    st.error("❌ No se pudieron detectar automáticamente las columnas de 'Código de Barras' y 'Nombre/Descripción' en el archivo.")
            except Exception as ex:
                st.error(f"Error procesando el archivo: {ex}")

    with tab_import_image:
        st.subheader("📸 Extraer Códigos y Nombres desde Imagen (Masivo / Lista / Factura)")
        st.markdown("Sube una foto que contenga productos y códigos. Se actualizarán los existentes y se agregarán los nuevos.")
        
        batch_img = st.file_uploader("Sube la imagen con los códigos", type=["png", "jpg", "jpeg", "webp"], key="batch_img_upload")
        
        if batch_img is not None:
            st.image(batch_img, caption="Imagen cargada", width=400)
            if st.button("🚀 Extraer, Actualizar y Agregar desde Imagen"):
                keys_to_try = []
                if paid_api_key:
                    keys_to_try.append((paid_api_key, "Versión de Pago"))
                if free_key_1:
                    keys_to_try.append((free_key_1, "Respaldo #1"))
                if free_key_2:
                    keys_to_try.append((free_key_2, "Respaldo #2"))
                
                if not keys_to_try:
                    st.error("❌ No hay claves API configuradas.")
                else:
                    extracted_list = []
                    success_batch = False
                    
                    prompt_batch = (
                        "Analiza esta imagen y extrae todos los productos y códigos de barras visibles. "
                        "Devuelve la información estrictamente en formato JSON con una lista de objetos bajo la clave 'productos', "
                        "donde cada objeto tenga 'codigo_barras' y 'nombre_producto'. "
                        "Respuesta JSON pura sin texto adicional."
                    )
                    
                    with st.spinner("Extrayendo códigos y nombres desde la imagen..."):
                        for api_k, label in keys_to_try:
                            try:
                                genai.configure(api_key=api_k)
                                model = genai.GenerativeModel('gemini-3.6-flash')
                                
                                batch_img.seek(0)
                                b_bytes = batch_img.read()
                                
                                resp_b = model.generate_content([
                                    {'mime_type': batch_img.type, 'data': b_bytes},
                                    prompt_batch
                                ])
                                
                                raw_tb = resp_b.text.strip()
                                if raw_tb.startswith("```json"):
                                    raw_tb = raw_tb[7:]
                                if raw_tb.endswith("```"):
                                    raw_tb = raw_tb[:-3]
                                    
                                res_json = json.loads(raw_tb.strip())
                                extracted_list = res_json.get("productos", [])
                                success_batch = True
                                break
                            except Exception as e:
                                continue
                    
                    if success_batch and extracted_list:
                        total_detectados = len(extracted_list)
                        procesados_ok = 0
                        nuevos_agregados = 0
                        actualizados = 0
                        no_procesados = 0
                        motivos_no_procesados_img = []

                        for idx, item in enumerate(extracted_list, start=1):
                            code_v = str(item.get("codigo_barras", "")).strip()
                            name_v = str(item.get("nombre_producto", "")).strip().upper()
                            
                            nombre_articulo = name_v if (name_v and name_v.lower() not in ["nan", "none", ""]) else f"Ítem #{idx}"
                            
                            if not code_v or code_v.lower() in ["nan", "none", ""]:
                                no_procesados += 1
                                motivos_no_procesados_img.append(f"Ítem #{idx} ➔ **Artículo:** *{nombre_articulo}* | **Motivo:** Código de barras faltante o no detectado.")
                                continue
                                
                            if not name_v or name_v.lower() in ["nan", "none", ""]:
                                no_procesados += 1
                                motivos_no_procesados_img.append(f"Ítem #{idx} ➔ **Código:** `{code_v}` | **Motivo:** Nombre de producto faltante.")
                                continue
                            
                            procesados_ok += 1
                            if code_v in st.session_state["barcode_memory"]:
                                if st.session_state["barcode_memory"][code_v] != name_v:
                                    st.session_state["barcode_memory"][code_v] = name_v
                                    actualizados += 1
                            else:
                                st.session_state["barcode_memory"][code_v] = name_v
                                nuevos_agregados += 1
                        
                        save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                        
                        st.success("🎯 **¡Procesamiento de imagen finalizado!**")
                        col_i1, col_i2, col_i3, col_i4, col_i5 = st.columns(5)
                        col_i1.metric("Detectados", total_detectados)
                        col_i2.metric("Procesados", procesados_ok)
                        col_i3.metric("Nuevos Agregados", nuevos_agregados)
                        col_i4.metric("Actualizados", actualizados)
                        col_i5.metric("No Procesados", no_procesados)

                        if motivos_no_procesados_img:
                            with st.expander(f"⚠️ Ver detalle de los {len(motivos_no_procesados_img)} elementos no procesados y su artículo"):
                                for motivo in motivos_no_procesados_img:
                                    st.markdown(f"- {motivo}")
                    else:
                        st.error("No se pudieron extraer códigos estructurados de la imagen.")
