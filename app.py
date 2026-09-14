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

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")

master_dict = {}
master_names = []
master_code_to_details = {}

def parse_master_dataframe(df_m):
    global master_dict, master_names, master_code_to_details
    master_dict = {}
    master_names = []
    master_code_to_details = {}
    cols = [str(c).lower() for c in df_m.columns]
    
    name_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_m.columns[0])
    code_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_m.columns[1])
    price_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'precio' in c), None)
    cost_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'costo' in c), None)
    stock_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'stock' in c), None)
    cat_col = next((df_m.columns[i] for i, c in enumerate(cols) if 'categor' in c), None)
    
    for _, row in df_m.iterrows():
        p_name = str(row[name_col]).strip().upper()
        raw_code = str(row[code_col]).strip()
        if raw_code.endswith('.0'):
            raw_code = raw_code[:-2]
        p_code = raw_code
        
        if p_name and p_name != "NAN" and p_code and p_code != "NAN":
            master_dict[p_name] = p_code
            master_names.append(p_name)
            
            master_code_to_details[p_code] = {
                "nombre": p_name,
                "precio": row[price_col] if price_col else "N/D",
                "costo": row[cost_col] if cost_col else "N/D",
                "stock": row[stock_col] if stock_col else "N/D",
                "categoria": row[cat_col] if cat_col else "General"
            }

if os.path.exists(SAVED_MASTER_FILE) and "master_loaded_once" not in st.session_state:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE, dtype=str)
        parse_master_dataframe(df_saved)
        st.sidebar.success(f"📂 Maestro anterior cargado: {len(master_dict)} productos.")
        st.session_state["master_loaded_once"] = True
    except Exception as e:
        print(f"Error cargando maestro previo: {e}")

master_file_uploaded = st.sidebar.file_uploader("Actualizar Maestro (Sube nuevo archivo)", type=["xlsx", "xls", "csv"], key="master_inv_file")

if master_file_uploaded is not None:
    try:
        if master_file_uploaded.name.endswith('.csv'):
            df_master = pd.read_csv(master_file_uploaded, dtype=str)
            df_master.to_excel(SAVED_MASTER_FILE, index=False)
        else:
            df_master = pd.read_excel(master_file_uploaded, dtype=str)
            with open(SAVED_MASTER_FILE, "wb") as f:
                f.write(master_file_uploaded.getbuffer())
        
        parse_master_dataframe(df_master)
        st.sidebar.success(f"✅ Nuevo maestro guardado: {len(master_dict)} productos.")
    except Exception as e:
        st.sidebar.error(f"Error al procesar el maestro: {e}")
elif os.path.exists(SAVED_MASTER_FILE) and not master_dict:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE, dtype=str)
        parse_master_dataframe(df_saved)
    except Exception:
        pass

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
# MOTOR DE INTELIGENCIA Y RANGOS DE EMPAREJAMIENTO SEGURO
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
    clean_desc_key = str(item_description).strip().upper()
    
    if "CORONA CERO" in clean_desc_key or "CERO 355" in clean_desc_key:
        return "750304423180", "Actualizado (Regla Maestra Inmediata Corona Cero)"

    if not master_dict:
        return str(original_code).strip(), "Sin Maestro Cargado"
    
    clean_orig_code = str(original_code).strip()
    if clean_orig_code.endswith('.0'):
        clean_orig_code = clean_orig_code[:-2]
    if not clean_orig_code or clean_orig_code.lower() in ["nan", "none", ""]:
        clean_orig_code = ""

    if clean_desc_key in st.session_state["product_overrides"]:
        return str(st.session_state["product_overrides"][clean_desc_key]).strip(), "Actualizado (Regla Guardada)"

    prov_tokens, norm_prov = clean_and_normalize(item_description)
    prov_volume = extract_volume_token(item_description)
    
    for m_name, m_code in master_dict.items():
        _, norm_m = clean_and_normalize(m_name)
        if m_name == norm_prov or m_name == clean_desc_key or norm_m == norm_prov:
            return str(m_code).strip(), "Actualizado (Exacto)"
            
    best_match_code = clean_orig_code
    max_matched_tiers = 0
    highest_score = 0.0

    prov_set = {t for t in prov_tokens if len(t) > 2 or t == "CERO"}

    for m_name in master_names:
        m_tokens, _ = clean_and_normalize(m_name)
        m_volume = extract_volume_token(m_name)
        master_set = {t for t in m_tokens if len(t) > 2 or t == "CERO"}

        if not prov_set or not master_set:
            continue

        matched_tiers = 0
        common_tokens = prov_set.intersection(master_set)
        if not common_tokens:
            continue
        matched_tiers += 1

        if prov_volume and m_volume:
            if prov_volume != m_volume:
                continue
        matched_tiers += 1

        union_tokens = prov_set.union(master_set)
        jaccard = len(common_tokens) / len(union_tokens)
        
        score = jaccard
        for t in common_tokens:
            if len(t) > 3 or t == "CERO":
                score += 0.25

        if matched_tiers >= 2 and score >= 0.45:
            if score > highest_score:
                highest_score = score
                max_matched_tiers = matched_tiers
                best_match_code = str(master_dict[m_name]).strip()

    if highest_score >= 0.50 and max_matched_tiers >= 2:
        return best_match_code, "Actualizado (IA por Rangos)"

    if not clean_orig_code:
        return "S/C (Sin Código)", "⚠️ Sin Código Original en Factura"

    return clean_orig_code, "⚠️ Conserva Código Original (Sin Rango Seguro)"

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
        "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (número de unidades o bultos según la factura), 'empaque' (unidades individuales por caja/bulto. Si la factura muestra 'UN' o 'PC' con empaque 1, asigna empaque 1), y 'costo_sin_itbis' (EL COSTO UNITARIO REAL POR CADA PIEZA INDIVIDUAL: si la línea muestra un Impuesto Neto total y la cantidad es N con empaque 1, el costo unitario es Impuesto_Neto / cantidad. Si es por cajas, divide Impuesto_Neto / (cantidad * empaque)). "
        "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre comercial del producto y su presentación o tamaño limpio (ej: 'CORONA CERO 355 ML', 'MAESTRO DOBEL DIAMANTE 700 ML', 'EVIAN 75 CL'), eliminando códigos internos, diagonales de empaque y textos redundantes. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
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

        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            
            with st.spinner("Analizando factura con auditoría de costos y rangos estrictos..."):
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
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match or "Conserva" in status_match or "Sin Código" in status_match:
                        unmatched_items.append((idx, desc, orig_code))

                    raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    
                    costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)

                    raw_pv = (costo * 1.25) * 1.18
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
                        "Precio Venta (M5)": precio_venta,
                        "Estado Maestro": status_match
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
                        str(item_dict["Código Barra POS"]),
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
    st.title("📂 Procesador por Lotes (Con Validación de Rangos)")
    st.markdown("Sube varias facturas. El sistema validará los ítems mediante rangos estrictos para evitar falsos positivos.")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        with st.expander("👁️ Vista Previa de los Archivos en Lote"):
            for f_item in uploaded_files:
                st.markdown(f"**Archivo:** `{f_item.name}`")
                if "image" in f_item.type or f_item.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                    st.image(Image.open(f_item), caption=f_item.name, width=300)
                    f_item.seek(0)

        if st.button("🚀 Procesar Lote y Validar con Maestro"):
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
                        st.warning(f"⚠️ Archivo omitido por estar duplicado en este lote: **{file.name}** (Doc: {num_doc}, Total: {total_doc_val})")
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
            
            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y filtraron **{duplicate_count} archivo(s) duplicado(s)** dentro de la selección actual.")

            if invoice_totals_summary:
                st.markdown("### 🏢 Proveedores Identificados y Totales por Factura")
                df_totales = pd.DataFrame(invoice_totals_summary)
                st.dataframe(df_totales, use_container_width=True, hide_index=True)
                
                t_sub = sum(x["Subtotal"] for x in invoice_totals_summary)
                t_itbis = sum(x["ITBIS"] for x in invoice_totals_summary)
                t_gen = sum(x["Total General"] for x in invoice_totals_summary)
                
                c_l1, c_l2, c_l3 = st.columns(3)
                c_l1.metric("Subtotal Acumulado Lote", f"RD$ {t_sub:,.2f}")
                c_l2.metric("ITBIS Acumulado Lote", f"RD$ {t_itbis:,.2f}")
                c_l3.metric("Total General Acumulado", f"RD$ {t_gen:,.2f}")

            if all_consolidated_items:
                st.markdown("---")
                st.markdown(f"### 📦 Consolidado de Ítems ({len(all_consolidated_items)} productos totales)")

                rows_preview = []
                unmatched_batch = []

                for idx, item in enumerate(all_consolidated_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match or "Conserva" in status_match or "Sin Código" in status_match:
                        unmatched_batch.append((idx, desc, orig_code))

                    raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    
                    costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)

                    raw_pv = (costo * 1.25) * 1.18
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
                        "Precio Venta (M5)": precio_venta,
                        "Estado Maestro": status_match
                    })

                if unmatched_batch:
                    st.warning(f"⚠️ **Atención en lote:** Hay {len(unmatched_batch)} producto(s) sin match automático:")
                    for u_idx, u_desc, u_code in unmatched_batch:
                        st.markdown(f"- *{u_desc}*")

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
                        str(item_dict["Código Barra POS"]),
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

# ==========================================
# MÓDULO 3: EXTRAER CÓDIGO DESDE IMAGEN
# ==========================================
elif modulo == "📸 Extraer Código desde Imagen":
    st.title("📸 Lector de Códigos y Productos (Maestro, Memoria e Internet)")
    st.markdown("Sube la foto del código de barras o producto. Si no está en tu Maestro, el sistema lo buscará en internet y lo registrará en memoria.")

    img_uploaded = st.file_uploader("Sube la imagen del producto (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg", "webp"], key="barcode_img_upload")

    if img_uploaded is not None:
        st.success(f"Imagen cargada: {img_uploaded.name}")
        
        col_prev1, col_prev2 = st.columns([1, 1])
        with col_prev1:
            image_obj = Image.open(img_uploaded)
            st.image(image_obj, caption="Imagen analizada", use_container_width=True)
            img_uploaded.seek(0)
            
        with col_prev2:
            if st.button("🔍 Escanear y Buscar (Maestro e Internet)"):
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
                        "Extrae con absoluta precisión únicamente el número de código de barras visible (por ejemplo, los números debajo de las barras). "
                        "Devuelve la respuesta estrictamente en formato JSON con esta estructura exacta: "
                        '{"codigo_barras": "..."}'
                        "Sin texto adicional, solo el JSON."
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
                                err_str = str(e)
                                if "429" in err_str or "Quota exceeded" in err_str:
                                    continue
                                else:
                                    st.error(f"Error con {label}: {e}")
                                    break
                    
                    if success and extracted_code:
                        st.markdown("### 🎯 Resultado del Escaneo")
                        st.info(f"🔢 **Código de Barras Detectado:** `{extracted_code}`")
                        
                        # 1. Buscar en el Excel Maestro
                        product_info = master_code_to_details.get(extracted_code, None)
                        
                        if product_info:
                            st.success("✅ ¡Encontrado directamente en tu Excel Maestro!")
                            st.markdown(f"🏷️ **Nombre Oficial:** **{product_info['nombre']}**")
                            st.markdown(f"💰 **Precio de Venta:** RD$ {product_info['precio']}")
                            st.markdown(f"📦 **Costo:** RD$ {product_info['costo']}")
                            st.markdown(f"📊 **Stock Actual:** {product_info['stock']}")
                        else:
                            # 2. Buscar en la memoria previa de códigos escaneados
                            if extracted_code in st.session_state["barcode_memory"]:
                                mem_name = st.session_state["barcode_memory"][extracted_code]
                                st.success("💾 ¡Encontrado en la memoria de códigos registrados!")
                                st.markdown(f"🏷️ **Nombre Registrado:** **{mem_name}**")
                            else:
                                # 3. Si no está en el maestro ni memoria, consultar en internet usando Gemini
                                st.warning("⚠️ Código nuevo (No está en tu maestro POS). Consultando coincidencias en Internet...")
                                
                                internet_product_name = "Producto Nuevo Escaneado"
                                with st.spinner("Buscando información del producto en internet..."):
                                    try:
                                        search_model = genai.GenerativeModel('gemini-3.6-flash')
                                        search_resp = search_model.generate_content(
                                            f"Identifica y da el nombre comercial exacto, marca y presentación en español del producto cuyo código de barras universal es: {extracted_code}. Responde únicamente con el nombre del producto, sin explicaciones."
                                        )
                                        if search_resp and search_resp.text:
                                            internet_product_name = search_resp.text.strip().upper()
                                    except Exception as ex:
                                        print(f"Error en búsqueda web: {ex}")
                                
                                # Registrar automáticamente en la memoria persistente
                                st.session_state["barcode_memory"][extracted_code] = internet_product_name
                                save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                                
                                st.success("✨ ¡Producto nuevo identificado mediante internet y registrado en tu memoria!")
                                st.markdown(f"🏷️ **Nombre Identificado:** **{internet_product_name}**")
                                st.markdown(f"📌 *El código `{extracted_code}` ha sido guardado automáticamente en tu archivo de memoria.*")
                    else:
                        st.error("No se pudo extraer un código de barras claro de la imagen.")

# ==========================================
# MÓDULO 4: VER CÓDIGOS ALMACENADOS (CON EXTRACCIÓN DESDE EXCEL O IMAGEN MASIVA)
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.title("📋 Listado de Códigos y Nombres Almacenados")
    st.markdown("Consulta y alimenta tu memoria de códigos escaneados mediante carga de Excel o lectura masiva de imágenes.")

    # Pestañas secundarias para organizar la carga y visualización
    tab_view, tab_import_excel, tab_import_image = st.tabs(["📊 Ver Almacenados", "📂 Extraer desde Excel", "📸 Leer desde Imagen"])

    with tab_view:
        b_mem = st.session_state["barcode_memory"]
        if b_mem:
            st.info(f"📊 Total de códigos registrados en memoria: **{len(b_mem)}**")
            list_data = [{"Código de Barras": code, "Nombre del Producto": name} for code, name in b_mem.items()]
            df_codes = pd.DataFrame(list_data)
            st.dataframe(df_codes, use_container_width=True, hide_index=True)

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
        st.subheader("📂 Importar Códigos y Nombres desde un Archivo Excel o CSV")
        st.markdown("Sube un archivo Excel o CSV que contenga columnas de códigos de barras y descripciones para alimentar la memoria de golpe.")
        
        excel_import_file = st.file_uploader("Sube tu archivo Excel o CSV", type=["xlsx", "xls", "csv"], key="import_memory_file")
        
        if excel_import_file is not None:
            try:
                if excel_import_file.name.endswith('.csv'):
                    df_imp = pd.read_csv(excel_import_file, dtype=str)
                else:
                    df_imp = pd.read_excel(excel_import_file, dtype=str)
                
                st.write("Vista previa del archivo cargado:", df_imp.head(3))
                
                cols_lower = [str(c).lower() for c in df_imp.columns]
                c_code = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'codigo' in c or 'barra' in c or 'barcode' in c), None)
                c_name = next((df_imp.columns[i] for i, c in enumerate(cols_lower) if 'nombre' in c or 'descripcion' in c), None)
                
                if c_code and c_name:
                    if st.button("📥 Importar y Combinar con Memoria Actual"):
                        added_count = 0
                        for _, row in df_imp.iterrows():
                            c_val = str(row[c_code]).strip()
                            n_val = str(row[c_name]).strip().upper()
                            if c_val and c_val != "NAN" and n_val and n_val != "NAN":
                                if c_val.endswith('.0'):
                                    c_val = c_val[:-2]
                                st.session_state["barcode_memory"][c_val] = n_val
                                added_count += 1
                        
                        save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                        st.success(f"✅ ¡Se han importado y guardado exitosamente **{added_count}** productos en la memoria!")
                        st.rerun()
                else:
                    st.error("❌ No se pudieron detectar automáticamente las columnas de 'Código de Barras' y 'Nombre/Descripción' en el archivo.")
            except Exception as ex:
                st.error(f"Error procesando el archivo: {ex}")

    with tab_import_image:
        st.subheader("📸 Extraer Códigos y Nombres desde Imagen (Masivo / Lista / Factura)")
        st.markdown("Sube una foto o factura que contenga varios productos con sus códigos de barras. La IA los extraerá y registrará todos en tu memoria.")
        
        batch_img = st.file_uploader("Sube la imagen con los códigos", type=["png", "jpg", "jpeg", "webp"], key="batch_img_upload")
        
        if batch_img is not None:
            st.image(batch_img, caption="Imagen cargada", width=400)
            if st.button("🚀 Extraer y Registrar Códigos de la Imagen"):
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
                        "Ejemplo: {'productos': [{'codigo_barras': '...', 'nombre_producto': '...'}]}. "
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
                                err_str = str(e)
                                if "429" in err_str or "Quota exceeded" in err_str:
                                    continue
                                else:
                                    st.error(f"Error con {label}: {e}")
                                    break
                    
                    if success_batch and extracted_list:
                        new_added = 0
                        for item in extracted_list:
                            code_v = str(item.get("codigo_barras", "")).strip()
                            name_v = str(item.get("nombre_producto", "")).strip().upper()
                            if code_v and code_v != "NAN" and name_v and name_v != "NAN":
                                st.session_state["barcode_memory"][code_v] = name_v
                                new_added += 1
                        
                        save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                        st.success(f"✨ ¡Se han extraído y registrado con éxito **{new_added}** productos desde la imagen en tu memoria!")
                        st.rerun()
                    else:
                        st.error("No se pudieron extraer códigos estructurados de la imagen.")
