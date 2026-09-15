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

# Archivo de persistencia de memoria
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

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    t = text.upper()
    t = t.replace(' 5CL', ' 50 ML').replace(' 5 CL', ' 50 ML').replace('5CL', '50 ML')
    t = t.replace(' 75CL', ' 750 ML').replace(' 75 CL', ' 750 ML').replace('75CL', '750 ML')
    t = t.replace(' 1L', ' 1000 ML').replace(' 1 LT', ' 1000 ML').replace('1L', '1000 ML')
    t = t.replace(' 33CL', ' 330 ML').replace('33CL', '330 ML').replace(' 50CL', ' 500 ML').replace('50CL', '500 ML')
    t = t.replace(' 3 LT', ' 3000 ML').replace(' 3LT', ' 3000 ML')
    
    t = t.replace('PTE.', 'PRESIDENTE').replace('HU', '').replace('CJ', '').replace('BOT.', '').replace('LATA', 'LATA')
    t = t.replace(' 120Z', ' 12 OZ')
    
    for ch in ['/', '-', ',', '.', '(', ')', '%', '+', '"', "'"]:
        t = t.replace(ch, ' ')
    return " ".join(t.split())

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
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"]:
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
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📋 Ver Códigos Almacenados"]
)

# ==========================================
# MOTOR MAESTRO INTELIGENTE (INDEPENDIENTE DEL ORDEN)
# ==========================================
def match_official_barcode(item_description):
    raw_name = str(item_description).strip().upper()
    b_mem = st.session_state["barcode_memory"]
    
    if not b_mem:
        return "S/C (Sin Código)", raw_name, "⚠️ Memoria Vacía"

    # 1. Coincidencia Exacta
    if raw_name in b_mem:
        return clean_barcode(b_mem[raw_name]), raw_name, "Maestro Exacto"

    norm_input = normalize_text(raw_name)
    input_tokens = set(norm_input.split())

    best_score = 0.0
    best_code = "S/C (Sin Código)"
    best_name = raw_name

    # 2. Búsqueda por solapamiento de palabras clave (tokens), sin importar el orden
    for master_name, code in b_mem.items():
        norm_master = normalize_text(master_name)
        master_tokens = set(norm_master.split())
        
        if not master_tokens:
            continue

        common_tokens = input_tokens.intersection(master_tokens)
        token_score = len(common_tokens) / max(len(input_tokens), len(master_tokens))
        seq_ratio = difflib.SequenceMatcher(None, norm_input, norm_master).ratio()
        
        combined_score = (token_score * 0.7) + (seq_ratio * 0.3)
        
        if combined_score > best_score:
            best_score = combined_score
            best_code = code
            best_name = master_name

    # Umbral flexible para garantizar cruce de nombres con distintas estructuras
    if best_score >= 0.40:
        return clean_barcode(best_code), best_name, f"Smart Match ({best_score:.2f})"

    return "S/C (Sin Código)", raw_name, "⚠️ Sin Coincidencia en Maestro"

def audit_and_correct_cost(costo_unit, cantidad, empaque):
    c = safe_float(costo_unit)
    cant = safe_int(cantidad, 1)
    emp = safe_int(empaque, 1)
    
    if emp <= 1:
        return c
        
    if c > 800 and (c / emp) < c:
        if (c / emp) >= 5:
            return c / emp
            
    return c

def process_invoice_with_ai(file_obj, file_type):
    prompt_text = (
        "Analiza esta factura detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'emisor_nombre', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae unícamente: 'descripcion', 'cantidad', 'empaque', y 'costo_sin_itbis'. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    keys_to_try = []
    if paid_api_key:
        keys_to_try.append((paid_api_key, "Paid Tier"))
    if free_key_1:
        keys_to_try.append((free_key_1, "Free Key 1"))
    if free_key_2:
        keys_to_try.append((free_key_2, "Free Key 2"))

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
            return parsed_data, f"✅ Éxito"
        except Exception:
            continue
    return None, ""

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura. Los códigos oficiales de tu maestro se asignarán automáticamente por nombre.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, c_col2 = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            with st.spinner("Analizando factura y asignando códigos oficiales del maestro..."):
                parsed_data, success_msg = process_invoice_with_ai(uploaded_file, file_type)

            if parsed_data:
                st.success(success_msg)
                data_items = parsed_data.get("items", [])
                rows_preview = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                for idx, item in enumerate(data_items, start=1):
                    desc = str(item.get("descripcion", ""))
                    official_code, matched_name, status_match = match_official_barcode(desc)

                    raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    
                    costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)
                    raw_pv = (costo * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = cant_comprada * empaque_val
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Oficial POS": str(official_code),
                        "Nombre Maestro / Artículo": matched_name,
                        "Cant. Compra": cant_comprada,
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unitario": costo,
                        "Precio Venta": precio_venta,
                        "Estado": status_match
                    })

                df_resultado = pd.DataFrame(rows_preview)
                df_resultado["Código Oficial POS"] = df_resultado["Código Oficial POS"].astype(str)
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                wb = openpyxl.Workbook()
                ws_prod = wb.active
                ws_prod.title = "Productos"
                ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                
                for item_dict in rows_preview:
                    row_cells = [
                        item_dict["Nombre Maestro / Artículo"],
                        str(item_dict["Código Oficial POS"]),
                        "General",
                        "producto",
                        item_dict["Precio Venta"],
                        item_dict["Costo Unitario"],
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
                st.download_button("📥 Descargar Excel WilPOS Oficial", output.getvalue(), "Inventario_WilPOS_Actualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes y Consolidación Oficial</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas asignando los códigos correctos del maestro y consolidando sin duplicados.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    l_col1, l_col2 = st.columns([1, 3])
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
            st.session_state["batch_signatures"] = set()
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["is_live_processing"] = False

        processed_so_far = st.session_state["batch_processed_count"]

        b_col1, b_col2 = st.columns(2)
        iniciar_btn = b_col1.button("🚀 Iniciar Lote y Consolidar", type="primary")
        reiniciar_lote = b_col2.button("🔄 Reiniciar Lote")

        if reiniciar_lote:
            st.session_state["batch_accumulated_items"] = []
            st.session_state["batch_audit_log"] = []
            st.session_state["batch_signatures"] = set()
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["is_live_processing"] = False
            if "cached_uploaded_files" in st.session_state:
                del st.session_state["cached_uploaded_files"]
            st.success("¡Lote reiniciado!")
            st.rerun()

        if iniciar_btn:
            st.session_state["is_live_processing"] = True
            st.rerun()

        is_live = st.session_state.get("is_live_processing", False)
        if is_live:
            if processed_so_far < total_files:
                file_info = cached_files[processed_so_far]
                current_num = processed_so_far + 1
                
                st.info(f"⚡ **Procesando {current_num} de {total_files}:** `{file_info['name']}`...")
                st.progress(processed_so_far / total_files)

                file_bytes_io = io.BytesIO(file_info["bytes"])
                parsed_data, err_msg = process_invoice_with_ai(file_bytes_io, file_info["type"])

                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    nombre_prov = str(parsed_data.get("emisor_nombre", "Desconocido")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc_val = safe_float(parsed_data.get("total", 0))
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc_val}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in st.session_state["batch_signatures"]:
                        st.session_state["batch_audit_log"].append({
                            "Archivo": file_info["name"], "Estado": "🔴 Omitido (Duplicado)"
                        })
                    else:
                        st.session_state["batch_signatures"].add(doc_signature)
                        st.session_state["batch_ok_count"] += 1
                        st.session_state["batch_audit_log"].append({
                            "Archivo": file_info["name"], "Estado": "🟢 OK"
                        })
                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            st.session_state["batch_accumulated_items"].extend(items)
                else:
                    st.session_state["batch_audit_log"].append({
                        "Archivo": file_info["name"], "Estado": "🔴 Error IA"
                    })

                st.session_state["batch_processed_count"] += 1
                st.rerun()
            else:
                st.session_state["is_live_processing"] = False
                st.success("🎉 ¡Lote finalizado!")
                st.rerun()

        if st.session_state["batch_processed_count"] > 0:
            st.markdown("---")
            st.markdown("## 📊 Consolidado de Inventario Resultante")

            raw_items = st.session_state["batch_accumulated_items"]
            multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)

            processed_rows = []
            for item in raw_items:
                desc = str(item.get("descripcion", ""))
                official_code, matched_name, _ = match_official_barcode(desc)

                raw_costo = safe_float(item.get("costo_sin_itbis", 0))
                cant_comprada = safe_int(item.get("cantidad", 1), 1)
                empaque_val = safe_int(item.get("empaque", 1), 1)

                costo = audit_and_correct_cost(raw_costo, cant_comprada, empaque_val)
                raw_pv = (costo * multiplicador_ganancia) * 1.18
                precio_venta = round_to_nearest_5(raw_pv)
                stock_val = cant_comprada * empaque_val

                processed_rows.append({
                    "Nombre": matched_name,
                    "Código Barra": str(official_code),
                    "Categoría": "General",
                    "Tipo": "producto",
                    "Precio Venta": precio_venta,
                    "Costo": costo,
                    "Stock": stock_val,
                    "Stock Mínimo": 5,
                    "ITBIS": 0.18,
                    "Unidad Medida": "unidad",
                    "Venta Granel": "No",
                    "Cantidad Empaque": empaque_val,
                    "Precio Variable": "No",
                    "Descuento %": 0,
                    "Descuento Monto": 0,
                    "Precio Especial": None,
                    "Descuento Activo": "No",
                    "Descuento Nota": None
                })

            df_temp = pd.DataFrame(processed_rows)

            df_grouped = df_temp.groupby(['Código Barra', 'Nombre'], as_index=False).agg({
                'Stock': 'sum',
                'Costo': 'mean',
                'Precio Venta': 'mean',
                'Categoría': 'first',
                'Tipo': 'first',
                'Stock Mínimo': 'first',
                'ITBIS': 'first',
                'Unidad Medida': 'first',
                'Venta Granel': 'first',
                'Cantidad Empaque': 'first',
                'Precio Variable': 'first',
                'Descuento %': 'first',
                'Descuento Monto': 'first',
                'Precio Especial': 'first',
                'Descuento Activo': 'first',
                'Descuento Nota': 'first'
            })

            df_final_preview = df_grouped.sort_values(by="Stock", ascending=False).reset_index(drop=True)
            df_final_preview["Código Barra"] = df_final_preview["Código Barra"].astype(str)

            st.success(f"✨ Se consolidaron **{len(df_final_preview)} productos únicos** con sus códigos oficiales asignados.")
            st.dataframe(df_final_preview, use_container_width=True, hide_index=True)

            wb = openpyxl.Workbook()
            ws_prod = wb.active
            ws_prod.title = "Productos"
            ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])

            for _, row in df_final_preview.iterrows():
                ws_prod.append([
                    row["Nombre"],
                    str(row["Código Barra"]),
                    row["Categoría"],
                    row["Tipo"],
                    row["Precio Venta"],
                    row["Costo"],
                    row["Stock"],
                    row["Stock Mínimo"],
                    row["ITBIS"],
                    row["Unidad Medida"],
                    row["Venta Granel"],
                    row["Cantidad Empaque"],
                    row["Precio Variable"],
                    row["Descuento %"],
                    row["Descuento Monto"],
                    row["Precio Especial"],
                    row["Descuento Activo"],
                    row["Descuento Nota"]
                ])
                ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'

            output = io.BytesIO()
            wb.save(output)
            st.download_button("📥 Descargar Excel Consolidado Final", output.getvalue(), "Inventario_WilPOS_Consolidado_Corregido.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# MÓDULO 3: VER CÓDIGOS ALMACENADOS & CARGAR MAESTRO
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.markdown("<h2>📋 Memoria de <span style='color: #0284c7;'>Códigos Almacenados</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu archivo Excel maestro y gestiona el diccionario oficial de productos.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📥 Cargar Archivo Maestro")
    st.markdown("Sube tu Excel Maestro (`.xlsx` o `.xls`) con columnas de códigos y nombres de productos.")
    
    master_upload = st.file_uploader("Sube tu Excel Maestro", type=["xlsx", "xls"], key="master_tab_uploader")
    if master_upload is not None:
        if st.button("💾 Procesar y Almacenar Maestro"):
            try:
                df_master = pd.read_excel(master_upload, sheet_name=0, dtype=str)
                new_memory = {}
                
                # Detección automática inteligente de columnas (Código vs Nombre)
                col_code = None
                col_name = None
                for c in df_master.columns:
                    c_low = str(c).lower()
                    if 'codigo' in c_low or 'barra' in c_low or 'barcode' in c_low:
                        col_code = c
                    elif 'nombre' in c_low or 'descripcion' in c_low or 'producto' in c_low:
                        col_name = c
                
                if not col_code or not col_name:
                    col_code = df_master.columns[0]
                    col_name = df_master.columns[1]

                for _, r in df_master.iterrows():
                    val_a = str(r[col_code]).strip()
                    val_b = str(r[col_name]).strip()
                    
                    # Asignación segura sin importar el orden de las columnas en el Excel
                    if val_a.isdigit() or len(val_a) <= 15:
                        c_val, n_val = val_a, val_b.upper()
                    else:
                        c_val, n_val = val_b, val_a.upper()

                    if c_val and n_val and c_val.lower() not in ["nan", "none", ""]:
                        if c_val.endswith('.0'):
                            c_val = c_val[:-2]
                        new_memory[n_val] = c_val
                
                if new_memory:
                    st.session_state["barcode_memory"] = new_memory
                    save_json_file(BARCODE_MEMORY_FILE, new_memory)
                    st.success(f"¡Se almacenaron {len(new_memory)} productos con éxito en el sistema!")
                    st.rerun()
                else:
                    st.error("No se encontraron registros válidos en el archivo.")
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.info(f"📊 Total de códigos oficiales almacenados: **{len(b_mem)}**")
        df_codes = pd.DataFrame([{"Código de Barras Oficial": str(code), "Nombre del Producto": name} for name, code in b_mem.items()])
        df_codes["Código de Barras Oficial"] = df_codes["Código de Barras Oficial"].astype(str)
        st.dataframe(df_codes, use_container_width=True, hide_index=True, height=450)
        
        if st.button("🗑️ Borrar Memoria Almacenada"):
            st.session_state["barcode_memory"] = {}
            if os.path.exists(BARCODE_MEMORY_FILE):
                os.remove(BARCODE_MEMORY_FILE)
            st.success("¡Memoria borrada con éxito!")
            st.rerun()
    else:
        st.warning("⚠️ La memoria está vacía. Utiliza el botón de carga superior para subir tu archivo Excel maestro y almacenarlo en el sistema.")
