io
import json
import os
import hashlib
import re
import difflib
from datetime import datetime
import google.generativeai as genai
from PIL import Image
import streamlit as st
import openpyxl
import pandas as pd

# Intento de importar herramientas de renderizado PDF para vista previa universal
try:
    from pdf2image import convert_from_bytes
    PDF_RENDER_AVAILABLE = True
except ImportError:
    PDF_RENDER_AVAILABLE = False

# Configuración de la página
st.set_page_config(page_title="WilPOS - Automatizador de Facturas", page_icon="📊", layout="wide")

# Configuración de Claves API desde secrets de Streamlit o variables de entorno
free_key_1 = st.secrets.get("GEMINI_API_KEY_1", os.environ.get("GEMINI_API_KEY_1", ""))
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", os.environ.get("GEMINI_API_KEY_2", ""))
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", os.environ.get("GEMINI_API_KEY_PAID", ""))

# Archivo persistente de memoria de proveedores
MEMORY_FILE = "proveedores_memoria.json"
SAVED_MASTER_FILE = "maestro_guardado.xlsx"
MASTER_META_FILE = "maestro_meta.json"

def load_provider_memory():
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_provider_memory(memory_dict):
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(memory_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando memoria: {e}")

if "provider_memory" not in st.session_state:
    st.session_state["provider_memory"] = load_provider_memory()

# Funciones auxiliares de cálculo y formato con precisión estricta
def safe_float(val, default=0.0):
    try:
        if isinstance(val, str):
            val = val.replace("RD$", "").replace(",", "").strip()
        return round(float(val), 4)
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
st.sidebar.title("🧠 Memoria de Proveedores")
with st.sidebar.expander("Ver Proveedores Aprendidos"):
    mem = st.session_state["provider_memory"]
    if mem:
        st.write(f"Total proveedores en memoria: {len(mem)}")
        for rnc, info in mem.items():
            st.markdown(f"- **RNC:** `{rnc}` ({info.get('nombre', 'Sin nombre')})")
        if st.button("🗑️ Limpiar Memoria de Proveedores"):
            st.session_state["provider_memory"] = {}
            if os.path.exists(MEMORY_FILE):
                os.remove(MEMORY_FILE)
            st.success("¡Memoria reseteada!")
            st.rerun()
    else:
        st.info("Aún no hay proveedores aprendidos. Se registrarán automáticamente al procesar facturas.")

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")

# Gestión de Persistencia del Archivo Maestro
master_file_uploaded = st.sidebar.file_uploader("Actualizar archivo Maestro (opcional)", type=["xlsx", "xls", "csv"], key="master_inv_file")

master_dict = {}
master_names = []
master_by_code = {}
master_upload_date_str = "No disponible"

# Guardar nuevo maestro si el usuario lo subió
if master_file_uploaded is not None:
    try:
        bytes_data = master_file_uploaded.getvalue()
        with open(SAVED_MASTER_FILE, "wb") as f:
            f.write(bytes_data)
        
        current_time_str = datetime.now().strftime("%d/%m/%Y %I:%M %p")
        meta_data = {"fecha": current_time_str, "nombre_archivo": master_file_uploaded.name}
        with open(MASTER_META_FILE, "w", encoding="utf-8") as mf:
            json.dump(meta_data, mf)
            
        st.sidebar.success(f"✅ ¡Nuevo maestro guardado con éxito!")
    except Exception as ex:
        st.sidebar.error(f"Error al guardar maestro: {ex}")

# Cargar maestro activo (del archivo guardado o predeterminado local)
target_master_path = None
if os.path.exists(SAVED_MASTER_FILE):
    target_master_path = SAVED_MASTER_FILE
    if os.path.exists(MASTER_META_FILE):
        try:
            with open(MASTER_META_FILE, "r", encoding="utf-8") as mf:
                m_info = json.load(mf)
                master_upload_date_str = f"{m_info.get('fecha', 'N/D')} ({m_info.get('nombre_archivo', 'Archivo guardado')})"
        except Exception:
            master_upload_date_str = "Archivo guardado en sistema"
else:
    default_master_path = "Inventario_Completo_2026-09-06.xlsx"
    if os.path.exists(default_master_path):
        target_master_path = default_master_path
        mod_time = os.path.getmtime(default_master_path)
        master_upload_date_str = f"{datetime.fromtimestamp(mod_time).strftime('%d/%m/%Y %I:%M %p')} (Predeterminado local)"

st.sidebar.markdown(f"🕒 **Maestro Activo Desde:**\n`{master_upload_date_str}`")

if target_master_path is not None:
    try:
        if target_master_path.endswith('.csv'):
            df_master = pd.read_csv(target_master_path)
        else:
            df_master = pd.read_excel(target_master_path, sheet_name='Productos' if target_master_path.endswith('.xlsx') else 0)
        
        cols = [str(c).lower() for c in df_master.columns]
        name_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_master.columns[0])
        code_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_master.columns[1])
        
        for _, row in df_master.iterrows():
            p_name = str(row[name_col]).strip().upper()
            p_code = str(row[code_col]).strip()
            if p_name and p_name != "NAN" and p_name != "NONE":
                master_dict[p_name] = p_code
                master_by_code[p_code] = p_name
                if p_name not in master_names:
                    master_names.append(p_name)
                
        st.sidebar.success(f"✅ Sincronizado: {len(master_dict)} productos en memoria.")
    except Exception as e:
        st.sidebar.error(f"Error al leer el maestro activo: {e}")
else:
    st.sidebar.warning("⚠️ No se detectó ningún archivo maestro cargado.")

# Diccionario de equivalencias personalizables guardado en session_state
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO"
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
# MOTOR DE NORMALIZACIÓN Y COMPARACIÓN CRUZADA DE PRESENTACIONES
# ==========================================
GLOBAL_SYNONYMS = {
    "JW ": "JOHNIE WALKER ",
    "JW.": "JOHNIE WALKER",
    "JOHNNIE": "JOHNIE",
    "BUCH ": "BUCHANANS ",
    "BUCHANAN": "BUCHANANS",
    "CHIV ": "CHIVAS REGAL ",
    "CHIVAS": "CHIVAS REGAL",
    "RON ": "",
    "WHISKY ": "",
    "WISKY ": "",
    "BOTELLA ": "",
    "BOTELL ": "",
    "CAJA ": ""
}

def normalize_text(text):
    upper = str(text).upper().strip()
    for k, v in st.session_state["custom_equivalences"].items():
        if k in upper:
            upper = upper.replace(k, v)
    for abbr, full in GLOBAL_SYNONYMS.items():
        upper = upper.replace(abbr, full)
    cleaned = re.sub(r'[^A-Z0-9\s]', ' ', upper)
    return re.sub(r'\s+', ' ', cleaned).strip()

def extract_and_normalize_size(text):
    match = re.search(r'(\d+(?:\.\d+)?\s*(?:ML|L|LT|CL|OZ))', str(text), re.IGNORECASE)
    if not match:
        return None
    raw = re.sub(r'\s+', '', match.group(1)).upper()
    if raw.endswith("LT") or (raw.endswith("L") and not raw.endswith("ML") and not raw.endswith("CL") and not raw.endswith("OZ")):
        num_str = raw.replace("LT", "").replace("L", "")
        try:
            val = float(num_str)
            ml_val = int(val * 1000)
            return f"{ml_val}ML"
        except ValueError:
            pass
    elif raw.endswith("CL"):
        try:
            val = float(raw.replace("CL", ""))
            return f"{int(val * 10)}ML"
        except ValueError:
            pass
    return raw

def get_tokens(text):
    norm = normalize_text(text)
    stopwords = {"DE", "EL", "LA", "LOS", "LAS", "Y", "EN", "UN", "UNA", "CON", "CMS", "CM", "ML", "L", "LT", "CL", "BOT", "SCATOLA"}
    tokens = [t for t in norm.split() if t not in stopwords and not t.isdigit()]
    return set(tokens), norm

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "⚠️ Sin Maestro Cargado"
    
    if original_code in master_by_code:
        return original_code, f"Actualizado (Código Directo: {master_by_code[original_code]})"

    norm_desc = normalize_text(item_description)
    if norm_desc in master_dict:
        return master_dict[norm_desc], "Actualizado (Exacto Normalizado)"
        
    inv_size = extract_and_normalize_size(item_description)
    desc_tokens, _ = get_tokens(item_description)
    
    best_match_code = original_code
    best_match_name = ""
    highest_score = 0.0
    
    for m_name, m_code in master_dict.items():
        m_size = extract_and_normalize_size(m_name)
        
        if inv_size and m_size and inv_size != m_size:
            continue
            
        m_tokens, _ = get_tokens(m_name)
        if not desc_tokens or not m_tokens:
            continue
            
        intersection = desc_tokens.intersection(m_tokens)
        union = desc_tokens.union(m_tokens)
        jaccard = len(intersection) / len(union) if union else 0
        
        weight = 1.0
        for token in intersection:
            if len(token) > 3:
                weight += 0.40
            else:
                weight += 0.10
                
        if inv_size and m_size and inv_size == m_size:
            weight += 0.50
            
        score = jaccard * weight
        
        if score > highest_score and len(intersection) >= 1:
            highest_score = score
            best_match_code = m_code
            best_match_name = m_name
            
    if highest_score >= 0.28:
        return best_match_code, f"Actualizado (IA Presentación Cruzada: {best_match_name})"
        
    matches = difflib.get_close_matches(norm_desc, list(master_dict.keys()), n=1, cutoff=0.30)
    if matches:
        matched_name = matches[0]
        m_size = extract_and_normalize_size(matched_name)
        if not inv_size or not m_size or inv_size == m_size:
            return master_dict[matched_name], f"Actualizado (Similitud: {matched_name})"

    return original_code, "⚠️ No Encontrado en Maestro (Presentación Única / Sin Coincidencia)"

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
        "Para cada ítem, extrae con extrema precisión: 'codigo' (código interno entre corchetes, ej: C071904), 'descripcion' (nombre completo del producto incluyendo su volumen exacto o presentación, ej: GREY GOOSE 50ml, VINO 750ml), 'cantidad' (número de unidades o cajas compradas), 'empaque' (unidades individuales por caja, por defecto 1), y 'costo_sin_itbis' (EL COSTO UNITARIO REAL EXACTO POR CADA PIEZA INDIVIDUAL: si la factura muestra el monto total de la línea, divídelo estrictamente entre Cantidad * Empaque para obtener el costo de una sola pieza suelta). "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "REGLA CRÍTICA: Respuesta JSON pura sin texto adicional."
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
        success_msg = "✅ ¡Factura procesada con éxito y formato aprendido!"
        
        rnc_key = str(parsed_data.get("emisor_rnc", "")).strip()
        nombre_prov = str(parsed_data.get("emisor_nombre", "Proveedor Desconocido")).strip()
        
        if rnc_key and rnc_key not in st.session_state["provider_memory"]:
            st.session_state["provider_memory"][rnc_key] = {
                "nombre": nombre_prov if nombre_prov and nombre_prov != "None" else f"Proveedor RNC {rnc_key}",
                "nota_formato": "Formato de cajas con empaques y costos unitarios por pieza procesados exitosamente."
            }
            save_provider_memory(st.session_state["provider_memory"])

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
# CONSOLIDACIÓN Y NOTIFICACIÓN DE CRUCES ENTRE FACTURAS
# ==========================================
def consolidate_items_with_tracking(raw_items_with_source):
    consolidated_dict = {}
    cross_notifications = []
    
    for entry in raw_items_with_source:
        invoice_name = entry["invoice_name"]
        item = entry["item"]
        
        desc = str(item.get("descripcion", ""))
        orig_code = str(item.get("codigo", "")).strip()
        final_code, status_match = validate_with_master(desc, orig_code)
        
        key = final_code if final_code and "No Encontrado" not in status_match else desc.upper()
        
        cant = safe_int(item.get("cantidad", 1), 1)
        empaque = safe_int(item.get("empaque", 1), 1)
        stock = cant * empaque
        costo = safe_float(item.get("costo_sin_itbis", 0))
        
        if key in consolidated_dict:
            existing = consolidated_dict[key]
            total_stock_prev = existing["stock_total"]
            new_total_stock = total_stock_prev + stock
            
            total_cost_spent = (existing["costo"] * total_stock_prev) + (costo * stock)
            avg_cost = total_cost_spent / new_total_stock if new_total_stock > 0 else costo
            
            existing["stock_total"] = new_total_stock
            existing["cantidad_comprada"] += cant
            existing["costo"] = round(avg_cost, 4)
            existing["sources"].append(invoice_name)
            
            cross_notifications.append({
                "producto": desc,
                "codigo": final_code,
                "factura": invoice_name,
                "stock_agregado": stock
            })
        else:
            consolidated_dict[key] = {
                "codigo_barra": final_code,
                "nombre": desc,
                "cantidad_comprada": cant,
                "empaque": empaque,
                "stock_total": stock,
                "costo": costo,
                "estado_maestro": status_match,
                "sources": [invoice_name]
            }
            
    return list(consolidated_dict.values()), cross_notifications

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.info(f"🗂️ **Maestro Activo en Uso:** `{master_upload_date_str}`")
    st.markdown("Sube tu factura para extraer sus ítems, validar presentaciones con tu maestro POS y generar la plantilla actualizada.")

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")

        file_type_check = uploaded_file.type if hasattr(uploaded_file, 'type') else ''
        is_img = "image" in file_type_check or uploaded_file.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp'))
        
        with st.expander(f"👁️ Vista Previa del Archivo: {uploaded_file.name}"):
            if is_img:
                image = Image.open(uploaded_file)
                st.image(image, caption=f"Vista previa: {uploaded_file.name}", use_container_width=True)
                uploaded_file.seek(0)
            elif uploaded_file.name.lower().endswith('.pdf') and PDF_RENDER_AVAILABLE:
                try:
                    uploaded_file.seek(0)
                    pdf_images = convert_from_bytes(uploaded_file.read(), first_page=1, last_page=1)
                    if pdf_images:
                        st.image(pdf_images[0], caption=f"Página 1 - {uploaded_file.name}", use_container_width=True)
                    uploaded_file.seek(0)
                except Exception as ex:
                    st.info(f"No se pudo renderizar la vista previa visual del PDF: {ex}")
            else:
                st.info(f"El archivo '{uploaded_file.name}' está cargado y listo para procesamiento.")

        if st.session_state["quota_exceeded"]:
            @st.dialog("⚠️ Confirmación Requerida: Límite de Cuota Alcanzado")
            def quota_modal():
                st.write("Se ha agotado la cuota de las cuentas gratuitas de Gemini (Error 429 / Quota Exceeded).")
                st.write("¿Deseas confirmar el uso de la versión de pago para procesar esta factura?")
                
                col_m1, col_m2 = st.columns(2)
                with col_m1:
                    if st.button("✅ Sí, Confirmar", type="primary"):
                        st.session_state["use_paid_now"] = True
                        st.session_state["quota_exceeded"] = False
                        st.rerun()
                with col_m2:
                    if st.button("❌ Cancelar"):
                        st.session_state["quota_exceeded"] = False
                        st.rerun()
            quota_modal()

        if st.button("🚀 Procesar Factura") or st.session_state["use_paid_now"]:
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            
            with st.spinner("Analizando factura y realizando comparación cruzada de presentaciones..."):
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
                matched_count = 0

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "No Encontrado" in status_match or "Sin Maestro" in status_match:
                        unmatched_items.append({
                            "No.": idx,
                            "Descripción Proveedor": desc,
                            "Código Original": orig_code,
                            "Estado": status_match
                        })
                    else:
                        matched_count += 1

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
                
                c_m1, c_m2 = st.columns(2)
                c_m1.metric("✅ Actualizados Exitosamente", f"{matched_count} ítems")
                c_m2.metric("⚠️ No Encontrados (Sin Match)", f"{len(unmatched_items)} ítems")

                if unmatched_items:
                    with st.expander(f"⚠️ Ver detalle de los {len(unmatched_items)} productos NO actualizados (requieren revisión o regla de equivalencia)"):
                        st.dataframe(pd.DataFrame(unmatched_items), use_container_width=True, hide_index=True)

                st.markdown("#### Tabla Completa Procesada")
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
    st.title("📂 Procesador por Lotes (Validación de Presentaciones y Trazabilidad)")
    st.info(f"🗂️ **Maestro Activo en Uso:** `{master_upload_date_str}`")
    st.markdown("Sube varias facturas. El sistema validará todas las presentaciones cruzadas, consolidará duplicados y te notificará los cruces.")

    if st.session_state["quota_exceeded"]:
        @st.dialog("⚠️ Confirmación Requerida: Límite de Cuota Alcanzado")
        def quota_modal_batch():
            st.write("Se ha agotado la cuota de las cuentas gratuitas de Gemini (Error 429 / Quota Exceeded).")
            st.write("¿Deseas confirmar el uso de la versión de pago para procesar este lote?")
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                if st.button("✅ Sí, Confirmar", type="primary"):
                    st.session_state["use_paid_now"] = True
                    st.session_state["quota_exceeded"] = False
                    st.rerun()
            with col_m2:
                if st.button("❌ Cancelar"):
                    st.session_state["quota_exceeded"] = False
                    st.rerun()
        quota_modal_batch()

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        st.markdown("### 👁️ Vista Previa Selectiva de Archivos")
        st.markdown("Haz clic en el botón de cualquier archivo para desplegar su vista previa:")
        
        for idx_f, f_item in enumerate(uploaded_files):
            with st.expander(f"👁️ [Ojito] Ver factura #{idx_f+1}: {f_item.name}"):
                f_type = f_item.type if hasattr(f_item, 'type') else ''
                if "image" in f_type or f_item.name.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                    st.image(Image.open(f_item), caption=f_item.name, width=500)
                    f_item.seek(0)
                elif f_item.name.lower().endswith('.pdf') and PDF_RENDER_AVAILABLE:
                    try:
                        f_item.seek(0)
                        pdf_images = convert_from_bytes(f_item.read(), first_page=1, last_page=1)
                        if pdf_images:
                            st.image(pdf_images[0], caption=f"Página 1 - {f_item.name}", use_container_width=True)
                        f_item.seek(0)
                    except Exception as ex:
                        st.info(f"No se pudo renderizar la vista previa visual del PDF: {ex}")
                else:
                    st.info(f"El archivo '{f_item.name}' está cargado correctamente.")

        run_batch_processing = st.button("🚀 Procesar Lote, Consolidar y Validar Presentaciones", type="primary") or st.session_state["use_paid_now"]

        if run_batch_processing:
            all_raw_items_with_source = []
            invoice_totals_summary = []
            duplicate_count = 0
            batch_signatures = set()

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}...")
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                
                file.seek(0)
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
                            for it in items:
                                all_raw_items_with_source.append({
                                    "invoice_name": file.name,
                                    "item": it
                                })

                progress_bar.progress((i + 1) / len(uploaded_files))

            if st.session_state["use_paid_now"]:
                st.session_state["use_paid_now"] = False

            status_text.text("Consolidando ítems, validando presentaciones y detectando cruces...")
            consolidated_items, cross_notifications = consolidate_items_with_tracking(all_raw_items_with_source)

            status_text.text("¡Procesamiento completo!")
            
            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y filtraron **{duplicate_count} archivo(s) duplicado(s)** dentro de la selección actual.")

            if cross_notifications:
                st.markdown("### 🔔 Notificación de Productos Cruzados (Múltiples Facturas)")
                st.info(f"Se detectaron **{len(cross_notifications)} coincidencias** de productos repetidos en distintas facturas del lote. Sus stocks se han sumado y sus costos se han promediado ponderadamente.")
                df_cross = pd.DataFrame(cross_notifications)
                df_cross.columns = ["Descripción del Producto", "Código Barra / Ref", "Factura de Cruce", "Stock Añadido"]
                st.dataframe(df_cross, use_container_width=True, hide_index=True)

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

            if consolidated_items:
                st.markdown("---")
                st.markdown(f"### 📦 Consolidado Único de Productos ({len(consolidated_items)} productos finales)")

                rows_preview = []
                unmatched_batch = []
                matched_batch_count = 0

                for idx, c_item in enumerate(consolidated_items, start=1):
                    desc = c_item["nombre"]
                    final_code = c_item["codigo_barra"]
                    status_match = c_item["estado_maestro"]
                    
                    if "No Encontrado" in status_match or "Sin Maestro" in status_match:
                        unmatched_batch.append({
                            "No.": idx,
                            "Descripción Proveedor": desc,
                            "Código Original": final_code,
                            "Estado": status_match
                        })
                    else:
                        matched_batch_count += 1

                    costo = c_item["costo"]
                    raw_pv = (costo * 1.25) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = c_item["stock_total"]
                    empaque_val = c_item["empaque"]
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Barra POS": final_code,
                        "Nombre": desc,
                        "Cant. Compra": c_item["cantidad_comprada"],
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unit. Sin ITBIS": costo,
                        "Precio Venta (M5)": precio_venta,
                        "Fuentes": ", ".join(set(c_item["sources"])),
                        "Estado Maestro": status_match
                    })

                c_b1, c_b2 = st.columns(2)
                c_b1.metric("✅ Ítems Actualizados con Maestro", f"{matched_batch_count} ítems")
                c_b2.metric("⚠️ Ítems No Encontrados (Sin Match)", f"{len(unmatched_batch)} ítems")

                if unmatched_batch:
                    with st.expander(f"⚠️ Ver detalle de los {len(unmatched_batch)} productos NO encontrados en el maestro"):
                        st.dataframe(pd.DataFrame(unmatched_batch), use_container_width=True, hide_index=True)

                st.markdown("#### Tabla Consolidada Completa")
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
