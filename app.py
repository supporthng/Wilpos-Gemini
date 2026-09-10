import io
import json
import os
import hashlib
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

# Funciones auxiliares de cálculo
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

if "use_paid_now" not in st.session_state:
    st.session_state["use_paid_now"] = False

# Ventana emergente (Modal centralizado de confirmación de cuota)
@st.dialog("⚠️ Confirmación Requerida: Límite de Cuota Alcanzado")
def paid_confirmation_dialog():
    st.write("Se ha agotado la cuota de las cuentas gratuitas de Gemini (Error 429 / Quota Exceeded).")
    st.write("¿Deseas confirmar el uso de la versión de pago para procesar esta factura?")
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        if st.button("✅ Sí, Confirmar", type="primary"):
            st.session_state["use_paid_now"] = True
            st.session_state["quota_exceeded"] = False
            st.rerun()
    with col_d2:
        if st.button("❌ Cancelar"):
            st.session_state["quota_exceeded"] = False
            st.rerun()

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL (MAESTRO Y EQUIVALENCIAS)
# ==========================================
st.sidebar.title("Menú de Navegación")
modulo = st.sidebar.radio(
    "Selecciona el Módulo",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")
master_file_uploaded = st.sidebar.file_uploader("Sube tu archivo Maestro", type=["xlsx", "xls", "csv"], key="master_inv_file")

master_dict = {}
master_names = []
if master_file_uploaded is not None:
    try:
        if master_file_uploaded.name.endswith('.csv'):
            df_master = pd.read_csv(master_file_uploaded)
        else:
            df_master = pd.read_excel(master_file_uploaded)
        
        cols = [c.lower() for c in df_master.columns]
        name_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_master.columns[0])
        code_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_master.columns[1])
        
        for _, row in df_master.iterrows():
            p_name = str(row[name_col]).strip().upper()
            p_code = str(row[code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
                
        st.sidebar.success(f"✅ Maestro cargado: {len(master_dict)} productos.")
    except Exception as e:
        st.sidebar.error(f"Error al leer el maestro: {e}")

# Diccionario de equivalencias personalizables guardado en session_state
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "EVIAN 24/0.5 LT": "EVIAN 500ML",
        "JW BLUE": "BLUE LABEL",
        "BUCHANANS 12": "BUCHANAN'S 12 AÑOS"
    }

st.sidebar.markdown("---")
st.sidebar.title("🔄 Equivalencias y Sinónimos")
st.sidebar.markdown("Mapea términos de proveedores a tu formato POS:")
with st.sidebar.expander("Ver / Editar Equivalencias"):
    eq_key = st.text_input("Término del Proveedor (ej: EVIAN 24/0.5 LT)")
    eq_val = st.text_input("Equivalente en tu POS (ej: EVIAN 500ML)")
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
# MOTOR DE NORMALIZACIÓN Y COMPARACIÓN ROBUSTA
# ==========================================
SYNONYMS_MAP = {
    "JW ": "JOHNNIE WALKER ",
    "JW.": "JOHNNIE WALKER",
    "BUCH ": "BUCHANANS ",
    "BUCHANAN": "BUCHANANS",
    "VODKA ABSOLUT": "ABSOLUT VODKA",
    "0.5 LT": "500ML",
    "0.5LT": "500ML",
    "1 LT": "1000ML",
    "1LT": "1000ML"
}

def normalize_product_name(name):
    clean = str(name).strip().upper()
    
    # 1. Aplicar reglas de equivalencia personalizadas del usuario
    for prov_term, pos_term in st.session_state["custom_equivalences"].items():
        if prov_term in clean:
            clean = clean.replace(prov_term, pos_term)
            
    # 2. Reemplazar abreviaturas comunes y conversiones de unidades
    for abbr, full in SYNONYMS_MAP.items():
        clean = clean.replace(abbr, full)
        
    return clean

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "Sin Maestro Cargado"
    
    norm_desc = normalize_product_name(item_description)
    
    # 1. Búsqueda exacta normalizada
    if norm_desc in master_dict:
        return master_dict[norm_desc], "Actualizado (Exacto Normalizado)"
    
    # 2. Búsqueda difusa estándar (Fuzzy matching)
    matches = difflib.get_close_matches(norm_desc, master_names, n=1, cutoff=0.50)
    if matches:
        matched_name = matches[0]
        return master_dict[matched_name], f"Actualizado (Similitud: {matched_name})"
    
    # 3. Búsqueda avanzada por tokens (palabras clave cruzadas)
    desc_tokens = set(norm_desc.split())
    best_match_code = original_code
    best_match_reason = "⚠️ No Encontrado en Maestro"
    max_shared_tokens = 0
    
    for m_name in master_names:
        m_tokens = set(m_name.split())
        shared = desc_tokens.intersection(m_tokens)
        if len(shared) >= 2 and len(shared) > max_shared_tokens:
            max_shared_tokens = len(shared)
            best_match_code = master_dict[m_name]
            best_match_reason = f"Actualizado (Tokens Clave: {m_name})"
            
    if max_shared_tokens >= 2:
        return best_match_code, best_match_reason

    return original_code, "⚠️ No Encontrado en Maestro"

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura para extraer sus ítems, validar códigos con tu maestro POS y generar la plantilla actualizada.")

    if st.session_state["quota_exceeded"]:
        paid_confirmation_dialog()

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")

        # 👁️ Botón de Vista Previa (Ojito)
        with st.expander("👁️ Vista Previa del Archivo Cargado"):
            file_type_check = uploaded_file.type if hasattr(uploaded_file, 'type') else ''
            if "image" in file_type_check or uploaded_file.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                image = Image.open(uploaded_file)
                st.image(image, caption=f"Vista previa: {uploaded_file.name}", use_container_width=True)
                uploaded_file.seek(0)
            else:
                st.info(f"El archivo '{uploaded_file.name}' es de tipo PDF o documento.")

        if st.button("🚀 Procesar Factura") or st.session_state["use_paid_now"]:
            with st.spinner("Analizando factura, validando maestro y calculando costos..."):
                prompt_text = (
                    "Analiza esta factura detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad', 'empaque', y 'costo_sin_itbis'. "
                    "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre del producto y su presentación o tamaño (ej: 'EVIAN 500ML' o 'BLUE LABEL 750ML'), eliminando códigos internos innecesarios o textos redundantes. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )
                
                file_bytes = uploaded_file.read()
                file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
                
                parsed_data = None
                success_msg = ""
                
                active_key = paid_api_key if st.session_state["use_paid_now"] else free_key_1
                
                try:
                    if not active_key and not st.session_state["use_paid_now"]:
                        active_key = free_key_2
                        
                    genai.configure(api_key=active_key if active_key else paid_api_key)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    
                    uploaded_file.seek(0)
                    file_bytes = uploaded_file.read()
                    
                    response = model.generate_content([
                        {'mime_type': file_type, 'data': file_bytes},
                        prompt_text
                    ])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    parsed_data = json.loads(raw_text)
                    st.session_state["use_paid_now"] = False
                    success_msg = "✅ ¡Factura procesada con éxito!"
                    
                except Exception as e:
                    err_str = str(e)
                    if ("429" in err_str or "Quota exceeded" in err_str) and not st.session_state["use_paid_now"]:
                        try:
                            genai.configure(api_key=free_key_2 if free_key_2 else paid_api_key)
                            model2 = genai.GenerativeModel('gemini-3.6-flash')
                            uploaded_file.seek(0)
                            response = model2.generate_content([
                                {'mime_type': file_type, 'data': uploaded_file.read()},
                                prompt_text
                            ])
                            raw_text = response.text.strip()
                            if raw_text.startswith("```json"):
                                raw_text = raw_text[7:]
                            if raw_text.endswith("```"):
                                raw_text = raw_text[:-3]
                            parsed_data = json.loads(raw_text.strip())
                            success_msg = "✅ ¡Factura procesada usando el respaldo gratuito #2!"
                        except Exception:
                            st.session_state["quota_exceeded"] = True
                            st.rerun()
                    else:
                        st.error(f"Error al procesar: {e}")

                if parsed_data:
                    st.success(success_msg)
                    
                    # Mostrar Totales de la Factura
                    st.markdown("### 📋 Resumen de Totales de la Factura")
                    c_t1, c_t2, c_t3 = st.columns(3)
                    c_t1.metric("Subtotal", f"RD$ {safe_float(parsed_data.get('subtotal', 0)):,.2f}")
                    c_t2.metric("ITBIS", f"RD$ {safe_float(parsed_data.get('itbis', 0)):,.2f}")
                    c_t3.metric("Total General", f"RD$ {safe_float(parsed_data.get('total', 0)):,.2f}")
                    
                    st.markdown("---")
                    st.markdown("### 📦 Validación con Maestro y Precios de Venta")

                    data_items = parsed_data.get("items", [])
                    rows_preview = []
                    unmatched_items = []

                    for idx, item in enumerate(data_items, start=1):
                        desc = str(item.get("descripcion", ""))
                        orig_code = str(item.get("codigo", "")).strip()
                        
                        final_code, status_match = validate_with_master(desc, orig_code)
                        if "No Encontrado" in status_match:
                            unmatched_items.append((desc, orig_code))

                        costo = safe_float(item.get("costo_sin_itbis", 0))
                        raw_pv = (costo * 1.25) * 1.18
                        precio_venta = round_to_nearest_5(raw_pv)
                        cant_comprada = safe_int(item.get("cantidad", 1), 1)
                        empaque_val = safe_int(item.get("empaque", 1), 1)
                        stock_val = cant_comprada * empaque_val
                        
                        rows_preview.append({
                            "No.": idx,
                            "Código Barra POS": final_code,
                            "Nombre": desc,
                            "Cant. Compra": cant_comprada,
                            "Empaque": empaque_val,
                            "Stock Total": stock_val,
                            "Costo Unit. Sin ITBIS": costo,
                            "Precio Venta (M5)": precio_venta,
                            "Estado Maestro": status_match
                        })
                    
                    if unmatched_items:
                        st.warning(f"⚠️ **Atención:** Hay {len(unmatched_items)} producto(s) que no se encontraron en tu archivo maestro y conservan su código original:")
                        for u_desc, u_code in unmatched_items:
                            st.markdown(f"- *{u_desc}* (Código original: `{u_code}`)")

                    df_resultado = pd.DataFrame(rows_preview)
                    st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                    
                    template_path = "Plantilla_Inventario_WilPOS_2.xlsx"
                    if not os.path.exists(template_path):
                        template_path = "Plantilla_Inventario_WilPOS.xlsx"
                        
                    if os.path.exists(template_path):
                        wb = openpyxl.load_workbook(template_path)
                        ws_prod = wb['Productos']
                        ws_prod.delete_rows(2, ws_prod.max_row)
                    else:
                        wb = openpyxl.Workbook()
                        ws_prod = wb.active
                        ws_prod.title = "Productos"
                        ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                    
                    for item_dict in rows_preview:
                        ws_prod.append([
                            item_dict["Nombre"],
                            item_dict["Código Barra POS"],
                            "General",
                            "producto",
                            item_dict["Precio Venta (M5)"],
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
    st.title("📂 Procesador por Lotes (Con Validación Avanzada de Maestro POS)")
    st.markdown("Sube varias facturas. El sistema filtrará duplicados y validará cada ítem contra tu maestro usando sinónimos y equivalencias.")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        # 👁️ Vista previa múltiple (Ojito para Lotes)
        with st.expander("👁️ Vista Previa de los Archivos en Lote"):
            for f_item in uploaded_files:
                st.markdown(f"**Archivo:** `{f_item.name}`")
                if "image" in f_item.type or f_item.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                    st.image(Image.open(f_item), caption=f_item.name, width=300)
                    f_item.seek(0)

        if st.button("🚀 Procesar Lote y Validar con Maestro", type="primary"):
            all_consolidated_items = []
            duplicate_count = 0
            batch_signatures = set()

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}...")

                prompt_text = (
                    "Analiza esta factura detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad', 'empaque', y 'costo_sin_itbis'. "
                    "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre del producto y su presentación o tamaño, eliminando códigos internos innecesarios o textos redundantes. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )

                file_bytes = file.read()
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                
                parsed_data = None
                
                try:
                    if free_key_1:
                        genai.configure(api_key=free_key_1)
                        model_b1 = genai.GenerativeModel('gemini-3.6-flash')
                        response = model_b1.generate_content([
                            {'mime_type': file_type, 'data': file_bytes},
                            prompt_text
                        ])
                        raw_txt = response.text.strip()
                        if raw_txt.startswith("```json"):
                            raw_txt = raw_txt[7:]
                        if raw_txt.endswith("```"):
                            raw_txt = raw_txt[:-3]
                        parsed_data = json.loads(raw_txt.strip())
                except Exception:
                    try:
                        if free_key_2:
                            genai.configure(api_key=free_key_2)
                            model_b2 = genai.GenerativeModel('gemini-3.6-flash')
                            response = model_b2.generate_content([
                                {'mime_type': file_type, 'data': file_bytes},
                                prompt_text
                            ])
                            raw_txt = response.text.strip()
                            if raw_txt.startswith("```json"):
                                raw_txt = raw_txt[7:]
                            if raw_txt.endswith("```"):
                                raw_txt = raw_txt[:-3]
                            parsed_data = json.loads(raw_txt.strip())
                    except Exception:
                        try:
                            if paid_api_key:
                                genai.configure(api_key=paid_api_key)
                                model_bp = genai.GenerativeModel('gemini-3.6-flash')
                                response = model_bp.generate_content([
                                    {'mime_type': file_type, 'data': file_bytes},
                                    prompt_text
                                ])
                                raw_txt = response.text.strip()
                                if raw_txt.startswith("```json"):
                                    raw_txt = raw_txt[7:]
                                if raw_txt.endswith("```"):
                                    raw_txt = raw_txt[:-3]
                                parsed_data = json.loads(raw_txt.strip())
                        except Exception as batch_err:
                            st.warning(f"No se pudo procesar el archivo {file.name}: {batch_err}")

                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in batch_signatures:
                        duplicate_count += 1
                        st.warning(f"⚠️ Archivo omitido por estar duplicado en este lote: **{file.name}** (Doc: {num_doc}, Total: {total_doc})")
                    else:
                        batch_signatures.add(doc_signature)
                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            all_consolidated_items.extend(items)

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.text("¡Procesamiento por lotes completado!")
            
            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y filtraron **{duplicate_count} archivo(s) duplicado(s)** dentro de la selección actual.")

            if all_consolidated_items:
                st.success(f"🎉 Se consolidaron exitosamente {len(all_consolidated_items)} ítems de facturas válidas.")

                rows_preview = []
                unmatched_batch = []

                for idx, item in enumerate(all_consolidated_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match:
                        unmatched_batch.append((desc, orig_code))

                    costo = safe_float(item.get("costo_sin_itbis", 0))
                    raw_pv = (costo * 1.25) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    stock_val = cant_comprada * empaque_val
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Barra POS": final_code,
                        "Nombre": desc,
                        "Cant. Compra": cant_comprada,
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unit. Sin ITBIS": costo,
                        "Precio Venta (M5)": precio_venta,
                        "Estado Maestro": status_match
                    })

                if unmatched_batch:
                    st.warning(f"⚠️ **Atención en lote:** Hay {len(unmatched_batch)} producto(s) no encontrados en el maestro:")
                    for u_desc, u_code in unmatched_batch:
                        st.markdown(f"- *{u_desc}* (Código original: `{u_code}`)")

                df_batch = pd.DataFrame(rows_preview)
                st.dataframe(df_batch, use_container_width=True, hide_index=True)

                template_path = "Plantilla_Inventario_WilPOS_2.xlsx"
                if not os.path.exists(template_path):
                    template_path = "Plantilla_Inventario_WilPOS.xlsx"
                    
                if os.path.exists(template_path):
                    wb = openpyxl.load_workbook(template_path)
                    ws_prod = wb['Productos']
                    ws_prod.delete_rows(2, ws_prod.max_row)
                else:
                    wb = openpyxl.Workbook()
                    ws_prod = wb.active
                    ws_prod.title = "Productos"
                    ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])

                for item_dict in rows_preview:
                    ws_prod.append([
                        item_dict["Nombre"],
                        item_dict["Código Barra POS"],
                        "General",
                        "producto",
                        item_dict["Precio Venta (M5)"],
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
            else:
                st.warning("No hay ítems válidos para consolidar.")
