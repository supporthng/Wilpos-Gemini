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
SAVED_MASTER_FILE = "ultimo_maestro_pos.xlsx"
TEMPLATE_FILE = "Plantilla_Inventario_WilPOS.xlsx"

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

# Funciones auxiliares de cálculo y formato
def safe_float(val, default=0.0):
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
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

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.title("Menú de Navegación")
modulo = st.sidebar.radio(
    "Selecciona el Módulo",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

st.sidebar.markdown("---")
st.sidebar.title("🧠 Memoria y Reglas POS")
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

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")

master_dict = {}
master_names = []

if os.path.exists(SAVED_MASTER_FILE) and "master_loaded_once" not in st.session_state:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE)
        cols_s = [c.lower() for c in df_saved.columns]
        s_name_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'nombre' in c or 'descripcion' in c), df_saved.columns[0])
        s_code_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_saved.columns[1])
        
        for _, row in df_saved.iterrows():
            p_name = str(row[s_name_col]).strip().upper()
            p_code = str(row[s_code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
        st.sidebar.success(f"📂 Maestro anterior cargado: {len(master_dict)} productos.")
        st.session_state["master_loaded_once"] = True
    except Exception as e:
        print(f"Error cargando maestro previo: {e}")

master_file_uploaded = st.sidebar.file_uploader("Actualizar Maestro (Sube nuevo archivo)", type=["xlsx", "xls", "csv"], key="master_inv_file")

if master_file_uploaded is not None:
    try:
        if master_file_uploaded.name.endswith('.csv'):
            df_master = pd.read_csv(master_file_uploaded)
            df_master.to_excel(SAVED_MASTER_FILE, index=False)
        else:
            df_master = pd.read_excel(master_file_uploaded)
            with open(SAVED_MASTER_FILE, "wb") as f:
                f.write(master_file_uploaded.getbuffer())
        
        master_dict = {}
        master_names = []
        cols = [c.lower() for c in df_master.columns]
        name_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_master.columns[0])
        code_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_master.columns[1])
        
        for _, row in df_master.iterrows():
            p_name = str(row[name_col]).strip().upper()
            p_code = str(row[code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
                
        st.sidebar.success(f"✅ Nuevo maestro guardado: {len(master_dict)} productos.")
    except Exception as e:
        st.sidebar.error(f"Error al procesar el maestro: {e}")
elif os.path.exists(SAVED_MASTER_FILE) and not master_dict:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE)
        cols_s = [c.lower() for c in df_saved.columns]
        s_name_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'nombre' in c or 'descripcion' in c), df_saved.columns[0])
        s_code_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_saved.columns[1])
        
        for _, row in df_saved.iterrows():
            p_name = str(row[s_name_col]).strip().upper()
            p_code = str(row[s_code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
    except Exception:
        pass

st.sidebar.markdown("---")
st.sidebar.title("📄 Plantilla Oficial WilPOS")
template_uploaded = st.sidebar.file_uploader("Actualizar Plantilla Base (.xlsx)", type=["xlsx"], key="template_file_uploader")
if template_uploaded is not None:
    with open(TEMPLATE_FILE, "wb") as f:
        f.write(template_uploaded.getbuffer())
    st.sidebar.success("✅ Plantilla oficial guardada.")

# Equivalencias personalizables
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL"
    }

# ==========================================
# MOTOR AUDITADO Y FLEXIBLE DE EMPAREJAMIENTO
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

def extract_volume_token(text):
    match = re.search(r'(\d+\s*(?:ML|L|CL))', str(text).upper())
    if match:
        v = match.group(1).replace(" ", "")
        if "L" in v and "ML" not in v:
            try:
                num = float(re.sub(r'[^0-9.]', '', v))
                return f"{int(num * 1000)}ML"
            except:
                pass
        return v
    return ""

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "Sin Maestro Cargado"
    
    clean_desc_key = item_description.strip().upper()
    
    # 1. Regla de Mapeo Manual / Sobrescrito (Prioridad Absoluta)
    if clean_desc_key in st.session_state["product_overrides"]:
        return st.session_state["product_overrides"][clean_desc_key], "Actualizado (Regla Guardada)"

    prov_tokens, norm_prov = clean_and_normalize(item_description)
    prov_volume = extract_volume_token(item_description)
    
    # 2. Búsqueda Exacta
    for m_name, m_code in master_dict.items():
        _, norm_m = clean_and_normalize(m_name)
        if m_name == norm_prov or m_name == clean_desc_key or norm_m == norm_prov:
            return m_code, "Actualizado (Exacto)"
            
    best_match_code = original_code
    highest_score = 0.0

    # Conjunto de tokens significativos del producto en la factura (sin importar el orden)
    prov_set = {t for t in prov_tokens if len(t) > 2 and not t.isdigit()}

    for m_name in master_names:
        m_tokens, _ = clean_and_normalize(m_name)
        m_volume = extract_volume_token(m_name)
        master_set = {t for t in m_tokens if len(t) > 2 and not t.isdigit()}

        if not prov_set or not master_set:
            continue

        # BLINDAJE FLEXIBLE DE PALABRAS CLAVE: Exigir que las palabras principales (ej: ENRIQUILLO y SODA) coexistan en ambos
        common_keywords = prov_set.intersection(master_set)
        if len(common_keywords) < min(len(prov_set), len(master_set)):
            # Si no comparten todas las palabras clave significativas, verificar solapamiento crítico
            if len(common_keywords) == 0:
                continue

        # BLINDAJE DE VOLUMEN: Si ambos tienen volumen especificado, deben coincidir
        if prov_volume and m_volume:
            if prov_volume != m_volume:
                continue

        # Cálculo de similitud Jaccard ponderado
        union = prov_set.union(master_set)
        score = len(common_keywords) / len(union)

        if score > highest_score and score >= 0.50:
            highest_score = score
            best_match_code = master_dict[m_name]

    if highest_score >= 0.50:
        return best_match_code, "Actualizado (IA Maestro Flexible)"

    return original_code, "⚠️ Conserva Código Original (Sin Match Seguro)"

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
        "Para cada ítem, extrae con absoluta precisión: 'codigo', 'descripcion', 'cantidad' (número de bultos/paquetes comprados), 'empaque' (unidades individuales que trae el paquete. REGLA ESPECIAL PARA ALOE PURE PLUS: si el código es 92713, el empaque es estrictamente 10 unidades por paquete. Para otros productos usa su empaque real como 24, 16, 6, o 1 si es unitario), y 'costo_sin_itbis'. "
        "REGLA CRÍTICA ABSOLUTA PARA EL COSTO UNITARIO: El costo devuelto en 'costo_sin_itbis' DEBE SER OBLIGATORIAMENTE EL COSTO POR UNIDAD SUELTA (PIEZA INDIVIDUAL), NUNCA EL COSTO DEL PAQUETE COMPLETO. "
        "Para calcularlo correctamente: toma el 'Imp. Neto' total de la línea y divídelo estrictamente entre (Cantidad de Paquetes × Empaque). Es decir: costo_sin_itbis = Imp. Neto / (Cantidad * Empaque). "
        "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre comercial y presentación limpia, eliminando códigos y textos redundantes. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
    )

    parsed_data = None
    success_msg = ""
    
    active_key = paid_api_key if st.session_state["use_paid_now"] else free_key_1
    if not active_key and not st.session_state["use_paid_now"]:
        active_key = free_key_2

    try:
        genai.configure(api_key=active_key if active_key else paid_api_key)
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
        success_msg = "✅ ¡Factura procesada con éxito y coincidencia flexible optimizada!"
        
        rnc_key = str(parsed_data.get("emisor_rnc", "")).strip()
        nombre_prov = str(parsed_data.get("emisor_nombre", "Proveedor Desconocido")).strip()
        
        if rnc_key and rnc_key not in st.session_state["provider_memory"]:
            st.session_state["provider_memory"][rnc_key] = {
                "nombre": nombre_prov if nombre_prov and nombre_prov != "None" else f"Proveedor RNC {rnc_key}",
                "nota_formato": "Formato procesado con coincidencia flexible de palabras clave."
            }
            save_json_file(MEMORY_FILE, st.session_state["provider_memory"])

        if st.session_state["use_paid_now"]:
            success_msg = "✅ ¡Factura procesada usando la Versión de Pago!"
            
    except Exception as e:
        err_str = str(e)
        if ("429" in err_str or "Quota exceeded" in err_str) and not st.session_state["use_paid_now"]:
            try:
                genai.configure(api_key=free_key_2 if free_key_2 else paid_api_key)
                model2 = genai.GenerativeModel('gemini-3.6-flash')
                file_obj.seek(0)
                response = model2.generate_content([
                    {'mime_type': file_type, 'data': file_obj.read()},
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

    return parsed_data, success_msg

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura para extraer sus ítems, validar códigos con tu maestro POS y generar la plantilla actualizada.")

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

        if st.button("🚀 Procesar Factura") or st.session_state["use_paid_now"]:
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            
            with st.spinner("Analizando factura con coincidencia flexible de palabras clave..."):
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
                st.markdown("### 📦 Validación con Maestro y Precios de Venta")

                data_items = parsed_data.get("items", [])
                rows_preview = []
                unmatched_items = []

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    if orig_code == "92713":
                        item["empaque"] = 10

                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match or "Conserva" in status_match:
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
                    st.warning(f"⚠️ **Atención:** Hay {len(unmatched_items)} producto(s) sin match seguro (se conservó su código original):")
                    for u_desc, u_code in unmatched_items:
                        st.markdown(f"- *{u_desc}* (Código original: `{u_code}`)")
                        with st.expander(f"➕ Asignar Código POS correcto para: {u_desc}"):
                            new_pos_code = st.text_input(f"Introduce el código POS correcto para '{u_desc}'", key=f"override_{u_desc}")
                            if st.button("Guardar Regla Mapeo", key=f"btn_override_{u_desc}"):
                                if new_pos_code:
                                    st.session_state["product_overrides"][u_desc.strip().upper()] = new_pos_code.strip()
                                    save_json_file(OVERRIDES_FILE, st.session_state["product_overrides"])
                                    st.success("¡Regla guardada con éxito! Vuelve a procesar para aplicarla.")
                                    st.rerun()

                df_resultado = pd.DataFrame(rows_preview)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                if os.path.exists(TEMPLATE_FILE):
                    wb = openpyxl.load_workbook(TEMPLATE_FILE)
                else:
                    wb = openpyxl.Workbook()
                    ws_default = wb.active
                    ws_default.title = "Productos"

                if "Productos" in wb.sheetnames:
                    ws_prod = wb["Productos"]
                else:
                    ws_prod = wb.create_sheet("Productos")

                if ws_prod.max_row > 1:
                    ws_prod.delete_rows(2, ws_prod.max_row)

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
                    label="📥 Descargar Excel Plantilla WilPOS Oficial Actualizada",
                    data=excel_data,
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.title("📂 Procesador por Lotes (Estructura Plantilla Oficial)")
    st.markdown("Sube varias facturas. Se consolidarán en una sola plantilla respetando todas las pestañas de tu archivo oficial.")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        if st.button("🚀 Procesar Lote y Consolidar", type="primary"):
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
                        st.warning(f"⚠️ Archivo duplicado omitido: **{file.name}**")
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

            if invoice_totals_summary:
                st.markdown("### 🏢 Totales por Factura")
                st.dataframe(pd.DataFrame(invoice_totals_summary), use_container_width=True, hide_index=True)

            if all_consolidated_items:
                st.markdown(f"### 📦 Consolidado Total ({len(all_consolidated_items)} ítems)")
                rows_preview = []
                for idx, item in enumerate(all_consolidated_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    if orig_code == "92713":
                        item["empaque"] = 10

                    final_code, status_match = validate_with_master(desc, orig_code)
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

                st.dataframe(pd.DataFrame(rows_preview), use_container_width=True, hide_index=True)

                if os.path.exists(TEMPLATE_FILE):
                    wb = openpyxl.load_workbook(TEMPLATE_FILE)
                else:
                    wb = openpyxl.Workbook()
                    ws_default = wb.active
                    ws_default.title = "Productos"

                if "Productos" in wb.sheetnames:
                    ws_prod = wb["Productos"]
                else:
                    ws_prod = wb.create_sheet("Productos")

                if ws_prod.max_row > 1:
                    ws_prod.delete_rows(2, ws_prod.max_row)

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
                    label="📥 Descargar Excel Consolidado Oficial WilPOS",
                    data=excel_data_batch,
                    file_name="Inventario_WilPOS_Consolidado_Oficial.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
