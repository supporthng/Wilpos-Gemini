import io
import json
import os
import time
import hashlib
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

# ------------------------------------------
# BÚSQUEDA UNIVERSAL DE CLAVES API
# ------------------------------------------
api_key_candidates = [
    st.secrets.get("GEMINI_API_KEY"),
    st.secrets.get("GOOGLE_API_KEY"),
    st.secrets.get("GEMINI_API_KEY_PAID"),
    st.secrets.get("GEMINI_API_KEY_1"),
    os.environ.get("GEMINI_API_KEY"),
    os.environ.get("GOOGLE_API_KEY"),
    os.environ.get("GEMINI_API_KEY_PAID"),
    os.environ.get("GEMINI_API_KEY_1")
]
ACTIVE_GEMINI_KEY = next((k for k in api_key_candidates if k and str(k).strip()), None)

openai_key_candidates = [
    st.secrets.get("OPENAI_API_KEY"),
    os.environ.get("OPENAI_API_KEY")
]
ACTIVE_OPENAI_KEY = next((k for k in openai_key_candidates if k and str(k).strip()), None)

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
    t = t.replace('SIX EIGHT NINE', '689').replace('SIX-EIGHT-NINE', '689')
    t = t.replace('COGÑA', 'COGNAC').replace('COGÑAC', 'COGNAC').replace('CONGNAC', 'COGNAC')
    t = t.replace('VSOP', 'V.S.O.P').replace('V S O P', 'V.S.O.P')
    t = t.replace('VS ', 'VERY SPECIAL ').replace(' VS', ' VERY SPECIAL')
    t = t.replace('GIN ', 'GINEBRA ').replace(' GIN', ' GINEBRA')
    t = t.replace(' 5CL', ' 50 ML').replace(' 5 CL', ' 50 ML').replace('5CL', '50 ML')
    t = t.replace(' 75CL', ' 750 ML').replace(' 75 CL', ' 750 ML').replace('75CL', '750 ML')
    t = t.replace(' 70CL', ' 700 ML').replace(' 70 CL', ' 700 ML').replace('70CL', '700 ML')
    t = t.replace('PTE.', 'PRESIDENTE').replace('HU', '').replace('CJ', '').replace('BOT.', '')
    for ch in ['/', '-', ',', '.', '(', ')', '%', '+', '"', "'", 'º']:
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
    return float(round(round(x / 5) * 5))

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
# MOTOR MAESTRO INTELIGENTE
# ==========================================
def match_official_barcode(item_description):
    raw_name = str(item_description).strip().upper()
    b_mem = st.session_state["barcode_memory"]
    
    if not b_mem:
        return "S/C (Sin Código)", raw_name, "⚠️ Memoria Vacía"

    if raw_name in b_mem:
        return clean_barcode(b_mem[raw_name]), raw_name, "Maestro Exacto"

    norm_input = normalize_text(raw_name)
    input_tokens = set(norm_input.split())

    best_score = 0.0
    best_code = "S/C (Sin Código)"
    best_name = raw_name

    for master_name, code in b_mem.items():
        norm_master = normalize_text(master_name)
        master_tokens = set(norm_master.split())
        
        if not master_tokens:
            continue

        common_tokens = input_tokens.intersection(master_tokens)
        token_score = len(common_tokens) / max(len(input_tokens), len(master_tokens))
        seq_ratio = difflib.SequenceMatcher(None, norm_input, norm_master).ratio()
        combined_score = (token_score * 0.65) + (seq_ratio * 0.35)
        
        if combined_score > best_score:
            best_score = combined_score
            best_code = code
            best_name = master_name

    if best_score >= 0.28:
        return clean_barcode(best_code), best_name, f"Smart Match ({best_score:.2f})"

    return "S/C (Sin Código)", raw_name, "⚠️ Sin Coincidencia"

def parse_empaque_from_tamano(tamano_txt, unidad_txt):
    u = str(unidad_txt).strip().upper()
    if "BOT" in u:
        return 1
        
    t = str(tamano_txt).strip()
    match_t = re.search(r'^(\d+)\s*/', t)
    if match_t:
        return int(match_t.group(1))
        
    return 12

def process_invoice_with_ai(file_obj, file_type, use_openai_fallback=False):
    if use_openai_fallback:
        if not ACTIVE_OPENAI_KEY:
            return None, "Falta clave API de OpenAI"
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"

        prompt_text = (
            "Analiza esta factura rigurosamente línea por línea de arriba a abajo, en el orden exacto del 1 al 18. "
            "Para cada renglón extrae estrictamente: 'descripcion', 'cantidad' (el número exacto de la columna CANTDAD), 'unidad' (lo que dice la columna UNID., CAJA o BOT.), "
            "'tamano' (lo que indica exactamente la columna TAMAÑO, ej: 12/75 CL. o 75 CL.), "
            "'precio_lista' (el precio exacto de la columna PRECIO para ese renglón), y 'descuento_porcentaje' (el porcentaje de la columna COM., ej: 10 para 10%). "
            "Devuelve un JSON puro con esta estructura exacta y en orden estricto: "
            '{"items": [{"descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
            "Respuesta JSON pura sin texto adicional ni markdown."
        )
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI)"
        except Exception as e:
            return None, str(e)

    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini"

    prompt_text = (
        "Analiza esta factura rigurosamente línea por línea de arriba a abajo, en el orden exacto del 1 al 18. "
        "Para cada renglón extrae estrictamente: 'descripcion', 'cantidad' (el número exacto de la columna CANTDAD), 'unidad' (lo que dice la columna UNID., CAJA o BOT.), "
        "'tamano' (lo que indica exactamente la columna TAMAÑO, ej: 12/75 CL. o 75 CL.), "
        "'precio_lista' (el precio exacto de la columna PRECIO para ese renglón), y 'descuento_porcentaje' (el porcentaje de la columna COM., ej: 10 para 10%). "
        "Devuelve un JSON puro con esta estructura exacta y en orden estricto: "
        '{"items": [{"descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito"
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
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura. Lectura exacta de empaques desde la columna TAMAÑO en orden original.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, _ = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    
    use_openai_single = st.checkbox("🤖 Usar OpenAI (GPT-4o-mini) para esta factura", value=False)
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        if st.button("🚀 Procesar Factura"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            with st.spinner("Analizando 18 renglones y empaques..."):
                parsed_data, success_msg = process_invoice_with_ai(uploaded_file, file_type, use_openai_fallback=use_openai_single)

            if success_msg == "QUOTA_EXCEEDED":
                st.error("⚠️ **Límite de cuota gratuita alcanzado en Gemini (Error 429).** Activa la casilla de OpenAI arriba.")
            elif not parsed_data or not isinstance(parsed_data, dict):
                st.error(f"⚠️ Error al procesar la factura con la IA: {success_msg}")
            else:
                st.success(success_msg)
                
                data_items = parsed_data.get("items", [])
                rows_preview = []
                omitted_items = []
                multiplicador_ganancia = 1 + (margen_ganancia / 100.0)

                calc_subtotal = 0.0
                calc_descuento_total = 0.0

                for idx, item in enumerate(data_items, start=1):
                    if not isinstance(item, dict):
                        omitted_items.append({"Item #": idx, "Descripción": str(item), "Razón": "Estructura inválida"})
                        continue
                    
                    desc = str(item.get("descripcion") or item.get("nombre") or "").strip()
                    if not desc:
                        omitted_items.append({"Item #": idx, "Descripción": "(Sin descripción)", "Razón": "Línea sin descripción"})
                        continue

                    precio_lista = safe_float(item.get("precio_lista") or 0)
                    if precio_lista <= 0:
                        omitted_items.append({"Item #": idx, "Descripción": desc, "Razón": "Precio de lista en 0"})
                        continue

                    official_code, matched_name, status_match = match_official_barcode(desc)
                    desc_pct = safe_float(item.get("descuento_porcentaje") or 0)
                    cant_comprada = safe_int(item.get("cantidad") or 1, 1)
                    unidad_txt = str(item.get("unidad") or "CAJA")
                    tamano_txt = str(item.get("tamano") or "12/75 CL.")
                    
                    empaque_val = parse_empaque_from_tamano(tamano_txt, unidad_txt)
                    
                    # Cálculos contables exactos
                    importe_bruto = precio_lista * cant_comprada
                    descuento_linea = importe_bruto * (desc_pct / 100.0)
                    importe_neto_linea = importe_bruto - descuento_linea
                    
                    calc_subtotal += importe_bruto
                    calc_descuento_total += descuento_linea

                    if empaque_val == 1:
                        total_unidades_linea = cant_comprada
                        costo = round(importe_neto_linea / cant_comprada, 2) if cant_comprada > 0 else round(importe_neto_linea, 2)
                    else:
                        total_unidades_linea = cant_comprada * empaque_val
                        costo = round(importe_neto_linea / total_unidades_linea, 2) if total_unidades_linea > 0 else round(importe_neto_linea, 2)

                    raw_pv = (costo * multiplicador_ganancia) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Oficial POS": str(official_code),
                        "Nombre Maestro / Artículo": matched_name,
                        "Cant. Compra": cant_comprada,
                        "Unidad": unidad_txt,
                        "Tamaño/Empaque": tamano_txt,
                        "Empaque Num": empaque_val,
                        "Stock Total": total_unidades_linea,
                        "Costo Unitario": costo,
                        "Precio Venta": precio_venta,
                        "Estado": status_match
                    })

                calc_neto_gravado = calc_subtotal - calc_descuento_total
                calc_itbis = calc_neto_gravado * 0.18
                calc_total_factura = calc_neto_gravado + calc_itbis

                st.markdown("### 📑 Totales Oficiales de la Factura (Calculados con Precisión en Python)")
                t1, t2, t3, t4 = st.columns(4)
                t1.metric("Subtotal Gravado", f"RD$ {calc_subtotal:,.2f}")
                t2.metric("Descuento Total", f"RD$ {calc_descuento_total:,.2f}")
                t3.metric("ITBIS Total (18%)", f"RD$ {calc_itbis:,.2f}")
                t4.metric("Total Neto Factura", f"RD$ {calc_total_factura:,.2f}")
                st.markdown("---")

                st.info(f"📋 **Auditoría de Lectura:** Se detectaron **{len(data_items)} ítems** brutos en la factura. Procesados con éxito: **{len(rows_preview)}** | Omitidos: **{len(omitted_items)}**")

                if rows_preview:
                    st.markdown("### ✅ Artículos Procesados Exitosamente (En orden original)")
                    df_resultado = pd.DataFrame(rows_preview)
                    df_resultado["Código Oficial POS"] = df_resultado["Código Oficial POS"].astype(str)
                    
                    altura_tabla = min(max(len(rows_preview) * 35 + 40, 200), 850)
                    st.dataframe(df_resultado, use_container_width=True, hide_index=True, height=altura_tabla)
                    
                    wb = openpyxl.Workbook()
                    ws_prod = wb.active
                    ws_prod.title = "Productos"
                    ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                    
                    for item_dict in rows_preview:
                        ws_prod.append([
                            item_dict["Nombre Maestro / Artículo"],
                            str(item_dict["Código Oficial POS"]),
                            "General", "producto",
                            item_dict["Precio Venta"],
                            item_dict["Costo Unitario"],
                            item_dict["Stock Total"],
                            5, 0.18, "unidad", "No",
                            item_dict["Empaque Num"], "No", 0, 0, None, "No", None
                        ])
                        ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                    
                    output = io.BytesIO()
                    wb.save(output)
                    st.download_button("📥 Descargar Excel WilPOS Oficial", output.getvalue(), "Inventario_WilPOS_Actualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                if omitted_items:
                    st.markdown("### ⚠️ Artículos Omitidos / Descartados")
                    st.dataframe(pd.DataFrame(omitted_items), use_container_width=True, hide_index=True, height=300)

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes y Consolidación Oficial</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas consolidando empaques y cantidades exactas.</p>", unsafe_allow_html=True)
    st.markdown("---")

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
            st.session_state["batch_signatures"] = set()
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["batch_omitted_summary"] = []
            st.session_state["is_live_processing"] = False
            st.session_state["quota_paused"] = False

        processed_so_far = st.session_state["batch_processed_count"]

        if st.session_state.get("quota_paused", False):
            st.error("⚠️ **Límite de cuota gratuita alcanzado.**")
            if st.button("🤖 Continuar con Respaldo OpenAI", type="primary"):
                st.session_state["quota_paused"] = False
                st.session_state["use_openai_batch"] = True
                st.session_state["is_live_processing"] = True
                st.rerun()
        else:
            b_col1, b_col2 = st.columns(2)
            iniciar_btn = b_col1.button("🚀 Iniciar Lote y Consolidar", type="primary")
            reiniciar_lote = b_col2.button("🔄 Reiniciar Lote")

            if reiniciar_lote:
                st.session_state["batch_accumulated_items"] = []
                st.session_state["batch_audit_log"] = []
                st.session_state["batch_signatures"] = set()
                st.session_state["batch_processed_count"] = 0
                st.session_state["batch_ok_count"] = 0
                st.session_state["batch_omitted_summary"] = []
                st.session_state["is_live_processing"] = False
                st.session_state["quota_paused"] = False
                st.session_state["use_openai_batch"] = False
                if "cached_uploaded_files" in st.session_state:
                    del st.session_state["cached_uploaded_files"]
                st.rerun()

            if iniciar_btn:
                st.session_state["use_openai_batch"] = False
                st.session_state["is_live_processing"] = True
                st.rerun()

        is_live = st.session_state.get("is_live_processing", False)
        use_openai_flag = st.session_state.get("use_openai_batch", False)

        if is_live and not st.session_state.get("quota_paused", False):
            if processed_so_far < total_files:
                file_info = cached_files[processed_so_far]
                st.info(f"⚡ Procesando {processed_so_far + 1} de {total_files}: `{file_info['name']}`...")
                st.progress(processed_so_far / total_files)

                file_bytes_io = io.BytesIO(file_info["bytes"])
                parsed_data, err_msg = process_invoice_with_ai(file_bytes_io, file_info["type"], use_openai_fallback=use_openai_flag)
                time.sleep(1.0)

                if err_msg == "QUOTA_EXCEEDED":
                    st.session_state["is_live_processing"] = False
                    st.session_state["quota_paused"] = True
                    st.rerun()

                if parsed_data and isinstance(parsed_data, dict):
                    st.session_state["batch_ok_count"] += 1
                    st.session_state["batch_audit_log"].append({"Archivo": file_info["name"], "Estado": "🟢 OK"})
                    items = parsed_data.get("items", [])
                    if isinstance(items, list):
                        for itm in items:
                            if isinstance(itm, dict):
                                d_txt = str(itm.get("descripcion") or itm.get("nombre") or "").strip()
                                p_val = safe_float(itm.get("precio_lista") or 0)
                                if d_txt and p_val > 0:
                                    st.session_state["batch_accumulated_items"].append(itm)
                                else:
                                    st.session_state["batch_omitted_summary"].append({"Archivo": file_info["name"], "Descripción": d_txt or "(Vacío)", "Razón": "Descripción vacía o precio 0"})
                else:
                    st.session_state["batch_audit_log"].append({"Archivo": file_info["name"], "Estado": f"🔴 Error: {err_msg}"})

                st.session_state["batch_processed_count"] += 1
                st.rerun()
            else:
                st.session_state["is_live_processing"] = False
                st.rerun()

        if st.session_state["batch_processed_count"] > 0:
            st.markdown("---")
            st.markdown("## 📊 Consolidado de Inventario y Totales Generales")

            raw_items = st.session_state["batch_accumulated_items"]
            multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)

            processed_rows = []
            total_unidades_inventario = 0

            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                
                desc = str(item.get("descripcion") or item.get("nombre") or "").strip()
                if not desc:
                    continue

                official_code, matched_name, _ = match_official_barcode(desc)
                precio_lista = safe_float(item.get("precio_lista") or 0)
                desc_pct = safe_float(item.get("descuento_porcentaje") or 0)
                cant_comprada = safe_int(item.get("cantidad") or 1, 1)
                unidad_txt = str(item.get("unidad") or "CAJA")
                tamano_txt = str(item.get("tamano") or "12/75 CL.")

                empaque_val = parse_empaque_from_tamano(tamano_txt, unidad_txt)
                
                importe_bruto = precio_lista * cant_comprada
                descuento_linea = importe_bruto * (desc_pct / 100.0)
                importe_neto_linea = importe_bruto - descuento_linea
                
                if empaque_val == 1:
                    total_unidades_linea = cant_comprada
                    costo = round(importe_neto_linea / cant_comprada, 2) if cant_comprada > 0 else round(importe_neto_linea, 2)
                else:
                    total_unidades_linea = cant_comprada * empaque_val
                    costo = round(importe_neto_linea / total_unidades_linea, 2) if total_unidades_linea > 0 else round(importe_neto_linea, 2)

                raw_pv = (costo * multiplicador_ganancia) * 1.18
                precio_venta = round_to_nearest_5(raw_pv)
                stock_val = total_unidades_linea
                total_unidades_inventario += stock_val

                processed_rows.append({
                    "Nombre": matched_name,
                    "Código Barra": str(official_code),
                    "Categoría": "General", "Tipo": "producto",
                    "Precio Venta": precio_venta, "Costo": costo,
                    "Stock": stock_val, "Stock Mínimo": 5, "ITBIS": 0.18,
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
                kpi2.metric("📦 Total Líneas Válidas", f"{len(raw_items)}")
                kpi3.metric("📦 Unidades en Stock", f"{total_unidades_inventario:,}")
                
                inversion_total_lote = (df_final_preview['Costo'] * df_final_preview['Stock']).sum()
                kpi4.metric("💰 Inversión Neta (Sin ITBIS)", f"RD$ {inversion_total_lote:,.2f}")

                st.markdown("---")
                
                altura_tabla_lote = min(max(len(df_final_preview) * 35 + 40, 200), 850)
                st.dataframe(df_final_preview, use_container_width=True, hide_index=True, height=altura_tabla_lote)

                if st.session_state.get("batch_omitted_summary"):
                    with st.expander("⚠️ Ver ítems omitidos en el lote"):
                        st.dataframe(pd.DataFrame(st.session_state["batch_omitted_summary"]), use_container_width=True, hide_index=True, height=300)

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
                st.download_button("📥 Descargar Excel Consolidado Final", output.getvalue(), "Inventario_WilPOS_Consolidado_Corregido.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("No se encontraron ítems válidos para consolidar.")

# ==========================================
# MÓDULO 3: MEMORIA
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.markdown("<h2>📋 Memoria de <span style='color: #0284c7;'>Códigos Almacenados</span></h2>", unsafe_allow_html=True)
    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.info(f"📊 Total de códigos oficiales almacenados: **{len(b_mem)}**")
        df_codes = pd.DataFrame([{"Código de Barras Oficial": str(code), "Nombre del Producto": name} for name, code in b_mem.items()])
        
        altura_tabla_mem = min(max(len(df_codes) * 35 + 40, 200), 850)
        st.dataframe(df_codes, use_container_width=True, hide_index=True, height=altura_tabla_mem)
    else:
        st.warning("⚠️ La memoria está vacía.")
