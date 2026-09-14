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
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS CSS SUAVES
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
        padding: 0.5rem 1.2rem;
        font-weight: 600;
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

# Equivalencias personalizables
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL",
        "CORONA CERO": "CERVEZA CORONA CERO"
    }

# ==========================================
# MOTOR DE INTELIGENCIA Y EMPAREJAMIENTO AUTOMÁTICO
# ==========================================
def validate_with_master(item_description, original_code):
    clean_desc_key = str(item_description).strip().upper()
    
    if "CORONA CERO" in clean_desc_key or "CERO 355" in clean_desc_key:
        return "750304423180", "Actualizado (Regla Maestra Inmediata Corona Cero)"

    clean_orig_code = str(original_code).strip()
    if clean_orig_code.endswith('.0'):
        clean_orig_code = clean_orig_code[:-2]
    if not clean_orig_code or clean_orig_code.lower() in ["nan", "none", ""]:
        clean_orig_code = ""

    # 1. Buscar en reglas manuales
    if clean_desc_key in st.session_state["product_overrides"]:
        return str(st.session_state["product_overrides"][clean_desc_key]).strip(), "Actualizado (Regla Guardada)"

    # 2. Búsqueda exacta en la memoria viva de códigos
    b_mem = st.session_state["barcode_memory"]
    for b_code, b_name in b_mem.items():
        if b_name == clean_desc_key:
            return str(b_code).strip(), "Actualizado (Memoria Viva Exacta)"

    # 3. Búsqueda por aproximación (Fuzzy Matching) en la memoria viva
    if b_mem:
        name_to_code = {str(name).upper(): code for code, name in b_mem.items()}
        close_matches = difflib.get_close_matches(clean_desc_key, name_to_code.keys(), n=1, cutoff=0.65)
        if close_matches:
            matched_name = close_matches[0]
            code_found = name_to_code[matched_name]
            return code_found, f"Actualizado (Memoria Viva por Similitud: '{matched_name}')"

    # 4. Si trae código original válido, lo aprendemos automáticamente
    if clean_orig_code and clean_orig_code != "S/C":
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
            if "429" in str(e) or "Quota exceeded" in str(e):
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
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura individual. El sistema consultará y actualizará automáticamente tu memoria viva de códigos.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, c_col2 = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input(
            "⚙️ Ganancia (%)", 
            min_value=0.0, 
            max_value=500.0, 
            value=25.0, 
            step=1.0, 
            key="textbox_individual"
        )
    
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")

        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            
            with st.spinner("Analizando factura y consultando memoria viva..."):
                parsed_data, success_msg = process_invoice_with_ai(uploaded_file, file_type)

            if parsed_data:
                st.success(success_msg)
                
                data_items = parsed_data.get("items", [])
                rows_preview = []
                unmatched_items = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    orig_code = str(item.get("codigo", "")).strip()
                    
                    final_code, status_match = validate_with_master(desc, orig_code)
                    if "Sin Código" in status_match:
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
                    st.warning(f"⚠️ **Atención:** Hay {len(unmatched_items)} producto(s) sin código detectado:")
                    for u_idx, u_desc, u_code in unmatched_items:
                        with st.expander(f"➕ Asignar Código POS para: {u_desc} (Ítem #{u_idx})"):
                            new_pos_code = st.text_input(f"Introduce el código POS correcto", key=f"override_{u_idx}_{u_code}")
                            if st.button("Guardar y Aprender", key=f"btn_override_{u_idx}_{u_code}"):
                                if new_pos_code:
                                    st.session_state["barcode_memory"][new_pos_code.strip()] = u_desc.strip().upper()
                                    save_json_file(BARCODE_MEMORY_FILE, st.session_state["barcode_memory"])
                                    st.success("¡Aprendido y guardado en memoria con éxito!")
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
                
                st.markdown("---")
                st.download_button(
                    label="📥 Descargar Excel Plantilla WilPOS Actualizada",
                    data=output.getvalue(),
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE) CON DETECCIÓN DE DUPLICADOS
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes con Antiduplicados</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas. El sistema filtrará automáticamente facturas duplicadas y alimentará tu memoria.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    l_col1, l_col2 = st.columns([1, 3])
    with l_col1:
        margen_ganancia_lote = st.number_input("⚙️ Ganancia (%) Lote", min_value=0.0, max_value=500.0, value=25.0, step=1.0)

    uploaded_files = st.file_uploader("📂 Sube tus facturas (Selección múltiple)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        if st.button("🚀 Procesar Lote, Filtrar Duplicados y Alimentar Memoria"):
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
                    total_doc_val = safe_float(parsed_data.get("total", 0))
                    
                    # Generar huella digital única para detectar duplicados
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc_val}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in batch_signatures:
                        duplicate_count += 1
                        st.warning(f"⚠️ Factura duplicada detectada y omitida: **{file.name}** (Documento Nº: {num_doc})")
                    else:
                        batch_signatures.add(doc_signature)
                        invoice_totals_summary.append({
                            "Proveedor": nombre_emisor if nombre_emisor and nombre_emisor != "None" else f"RNC: {rnc_emisor}",
                            "Archivo": file.name,
                            "Nº Documento": num_doc if num_doc else "N/D",
                            "Total General": total_doc_val
                        })

                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            all_consolidated_items.extend(items)

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.text("¡Procesamiento por lotes y filtro de duplicados completado!")

            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y filtraron **{duplicate_count} factura(s) duplicada(s)** en este lote.")
            else:
                st.success("✅ No se encontraron facturas duplicadas en el lote.")

            if all_consolidated_items:
                st.markdown("---")
                st.markdown(f"### 📦 Consolidado de Ítems ({len(all_consolidated_items)} productos totales)")

                rows_preview = []
                multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)

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

                st.markdown("---")
                st.download_button(
                    label="📥 Descargar Excel Consolidado Sin Duplicados",
                    data=output.getvalue(),
                    file_name="Inventario_WilPOS_Consolidado_Lote.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MÓDULO 3: EXTRAER CÓDIGO DESDE IMAGEN
# ==========================================
elif modulo == "📸 Extraer Código desde Imagen":
    st.markdown("<h2>📸 Lector de Códigos y <span style='color: #0284c7;'>Productos</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube una foto del producto. Si es nuevo, se registrará automáticamente en tu memoria.</p>", unsafe_allow_html=True)
    st.markdown("---")

    img_uploaded = st.file_uploader("📂 Sube la imagen del producto", type=["png", "jpg", "jpeg", "webp"], key="barcode_img_upload")
    if img_uploaded is not None:
        st.image(img_uploaded, caption="Imagen analizada", width=400)
        if st.button("🔍 Escanear y Registrar"):
            st.info("Escaneo rápido disponible en la versión completa.")

# ==========================================
# MÓDULO 4: VER CÓDIGOS ALMACENADOS
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.markdown("<h2>📋 Memoria Viva de <span style='color: #0284c7;'>Códigos Almacenados</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Esta es tu base de datos autogestionada. El sistema aprende y guarda aquí cada producto automáticamente.</p>", unsafe_allow_html=True)
    st.markdown("---")

    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.info(f"📊 Total de códigos aprendidos y almacenados: **{len(b_mem)}**")
        df_codes = pd.DataFrame([{"Código de Barras": code, "Nombre del Producto": name} for code, name in b_mem.items()])
        st.dataframe(df_codes, use_container_width=True, hide_index=True, height=450)

        output_db = io.BytesIO()
        with pd.ExcelWriter(output_db, engine='openpyxl') as writer:
            df_codes.to_excel(writer, index=False, sheet_name="Codigos_Almacenados")

        st.download_button(
            label="📥 Descargar Tu Memoria Actualizada (.xlsx)",
            data=output_db.getvalue(),
            file_name="codigos_barras_almacenados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ La memoria está vacía. A medida que proceses facturas, los códigos se irán acumulando aquí automáticamente.")
