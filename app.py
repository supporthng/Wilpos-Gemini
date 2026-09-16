import io
import json
import os
import time
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
gemini_paid_candidates = [
    st.secrets.get("GEMINI_API_KEY_PAID"),
    os.environ.get("GEMINI_API_KEY_PAID")
]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [
    st.secrets.get("GEMINI_API_KEY"),
    st.secrets.get("GOOGLE_API_KEY"),
    st.secrets.get("GEMINI_API_KEY_1"),
    os.environ.get("GEMINI_API_KEY"),
    os.environ.get("GOOGLE_API_KEY"),
    os.environ.get("GEMINI_API_KEY_1")
]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

openai_key_candidates = [
    st.secrets.get("OPENAI_API_KEY"),
    os.environ.get("OPENAI_API_KEY")
]
ACTIVE_OPENAI_KEY = next((k for k in openai_key_candidates if k and str(k).strip()), None)

BARCODE_MEMORY_FILE = "codigos_escaneados_memoria.json"

def load_json_file(filepath):
    data = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    
    # Códigos oficiales fijos verificados en factura
    data["WHISKY ESCOCES MALTA 12 AÑOS GLEN GRANT"] = "8000040630269"
    data["VINO TINTO SIX EIGHT NINE 689"] = "051497322618"
    data["VODKA SKYY"] = "721059007504"
    data["VODKA INFUSIONS CITRUS SKYY"] = "721059627504"
    data["VODKA INFUSIONS RASPBERRY SKYY"] = "721059637503"
    return data

if "barcode_memory" not in st.session_state:
    st.session_state["barcode_memory"] = load_json_file(BARCODE_MEMORY_FILE)

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = {}

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

def get_resolved_barcode(extracted_code, description):
    cleaned = clean_barcode(extracted_code)
    desc_upper = str(description).strip().upper()
    
    # 1. Si el código extraído es válido, lo usamos y actualizamos la memoria
    if cleaned != "S/C (Sin Código)":
        st.session_state["barcode_memory"][desc_upper] = cleaned
        return cleaned
    
    # 2. Si no viene en la factura, buscamos en el Catálogo Maestro cargado
    if desc_upper in st.session_state["master_catalog"]:
        return st.session_state["master_catalog"][desc_upper]
        
    # 3. Si no está en el maestro, buscamos en la memoria interna
    if desc_upper in st.session_state["barcode_memory"]:
        return st.session_state["barcode_memory"][desc_upper]
        
    return "S/C (Sin Código)"

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Automatizador Inteligente</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)", "📁 Cargar Archivo Maestro", "📋 Ver Códigos Almacenados"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Configuración de API")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

def parse_empaque_from_tamano(tamano_txt, unidad_txt):
    u = str(unidad_txt).strip().upper()
    if "BOT" in u:
        return 1
    t = str(tamano_txt).strip()
    match_t = re.search(r'^(\d+)\s*/', t)
    if match_t:
        return int(match_t.group(1))
    return 12

def process_invoice_exact_18(file_obj, file_type, use_paid_gemini=False, use_openai_fallback=False):
    prompt_text = (
        "Analiza esta factura de Álvarez & Sánchez con extrema precisión horizontal. La tabla tiene filas numeradas. "
        "Debes asegurar que el 'codigo_barras' de la columna izquierda esté estrictamente alineado con la 'descripcion' de esa misma línea horizontal exacta, sin desplazar los códigos hacia arriba ni hacia abajo. "
        "Para cada renglón extrae exactamente: "
        "1. 'codigo_barras': el número exacto de la columna 'CODIGO DE BARRAS' que se encuentra en la misma fila horizontal de la descripción. "
        "2. 'descripcion': el texto exacto de la columna 'DESCRIPCION'. "
        "3. 'cantidad': número de la columna 'CANTDAD'. "
        "4. 'unidad': 'CAJA' o 'BOT.'. "
        "5. 'tamano': texto exacto de la columna 'TAMAÑO' (ej: 12/75 CL., 6/70 CL., 75 CL.). "
        "6. 'precio_lista': número exacto de la columna 'PRECIO'. "
        "7. 'descuento_porcentaje': porcentaje de la columna 'COM.' (ej: 10). "
        "Devuelve un JSON puro con un arreglo exacto de objetos bajo la clave 'items': "
        '{"items": [{"codigo_barras": "...", "descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini:
        active_key = ACTIVE_GEMINI_FREE_KEY

    if not active_key and use_openai_fallback:
        use_openai_fallback = True

    if use_openai_fallback and not active_key:
        if not ACTIVE_OPENAI_KEY:
            return None, "Faltan claves API válidas (Gemini / OpenAI)"
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI gpt-4o)"
        except Exception as e:
            return None, str(e)

    for intento in range(2):
        try:
            genai.configure(api_key=active_key)
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
                if not use_paid_gemini and ACTIVE_GEMINI_PAID_KEY:
                    active_key = ACTIVE_GEMINI_PAID_KEY
                    continue
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
            
    if use_openai_fallback and ACTIVE_OPENAI_KEY:
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=4000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI Respaldo)"
        except Exception as e:
            return None, str(e)

    return None, last_err

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Extracción directa de códigos de barras oficiales impresos sin duplicados.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, _ = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    
    use_openai_single = st.checkbox("🤖 Permitir OpenAI como respaldo final si falla Gemini", value=False)
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        if st.button("🚀 Procesar Factura Exacta (18 Renglones)"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            with st.spinner("Leyendo códigos de barras y renglones de la factura..."):
                parsed_data, success_msg = process_invoice_exact_18(uploaded_file, file_type, use_paid_gemini=use_gemini_paid_api, use_openai_fallback=use_openai_single)

            if success_msg == "QUOTA_EXCEEDED":
                st.error("⚠️ **Límite de cuota alcanzado.** Activa la casilla de 'Gemini Paid (API de Pago)' en la barra lateral o habilita el respaldo de OpenAI.")
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

                    extracted_code = item.get("codigo_barras")
                    resolved_code = get_resolved_barcode(extracted_code, desc)

                    desc_pct = safe_float(item.get("descuento_porcentaje") or 0)
                    cant_comprada = safe_int(item.get("cantidad") or 1, 1)
                    unidad_txt = str(item.get("unidad") or "CAJA")
                    tamano_txt = str(item.get("tamano") or "12/75 CL.")
                    
                    empaque_val = parse_empaque_from_tamano(tamano_txt, unidad_txt)
                    
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
                        "Código Oficial POS": str(resolved_code),
                        "Nombre Maestro / Artículo": desc,
                        "Cant. Compra": cant_comprada,
                        "Unidad": unidad_txt,
                        "Tamaño/Empaque": tamano_txt,
                        "Empaque Num": empaque_val,
                        "Stock Total": total_unidades_linea,
                        "Costo Unitario": costo,
                        "Precio Venta": precio_venta,
                        "Estado": "Catálogo Maestro" if resolved_code != clean_barcode(extracted_code) else "Directo Factura"
                    })

                calc_neto_gravado = calc_subtotal - calc_descuento_total
                calc_itbis = calc_neto_gravado * 0.18
                calc_total_factura = calc_neto_gravado + calc_itbis

                st.markdown("### 📑 Totales Oficiales de la Factura")
                t1, t2, t3, t4 = st.columns(4)
                t1.metric("Subtotal Gravado", f"RD$ {calc_subtotal:,.2f}")
                t2.metric("Descuento Total", f"RD$ {calc_descuento_total:,.2f}")
                t3.metric("ITBIS Total (18%)", f"RD$ {calc_itbis:,.2f}")
                t4.metric("Total Neto Factura", f"RD$ {calc_total_factura:,.2f}")
                st.markdown("---")

                st.info(f"📋 **Auditoría de Lectura:** Se detectaron **{len(data_items)} ítems**. Procesados con éxito: **{len(rows_preview)}** | Omitidos: **{len(omitted_items)}**")

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

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes y Consolidación Oficial</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas consolidando empaques y códigos exactos.</p>", unsafe_allow_html=True)
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
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["is_live_processing"] = False
            st.session_state["quota_paused"] = False

        processed_so_far = st.session_state["batch_processed_count"]

        if st.session_state.get("quota_paused", False):
            st.error("⚠️ **Límite de cuota alcanzado.**")
            if st.button("💎 Reintentar con Gemini Paid o Activar Respaldo", type="primary"):
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
                parsed_data, err_msg = process_invoice_exact_18(file_bytes_io, file_info["type"], use_paid_gemini=use_gemini_paid_api, use_openai_fallback=False)
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
                                d_txt = str(itm.get("descripcion") or "").strip()
                                p_val = safe_float(itm.get("precio_lista") or 0)
                                if d_txt and p_val > 0:
                                    st.session_state["batch_accumulated_items"].append(itm)

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
                
                desc = str(item.get("descripcion") or "").strip()
                if not desc:
                    continue

                extracted_code = item.get("codigo_barras")
                resolved_code = get_resolved_barcode(extracted_code, desc)
                
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
                    "Nombre": desc,
                    "Código Barra": str(resolved_code),
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

# ==========================================
# MÓDULO 3: CARGAR ARCHIVO MAESTRO
# ==========================================
elif modulo == "📁 Cargar Archivo Maestro":
    st.markdown("<h2>📁 Gestión de <span style='color: #0284c7;'>Catálogo Maestro</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu archivo Excel maestro para mapear automáticamente los códigos de barras de los productos.</p>", unsafe_allow_html=True)
    st.markdown("---")

    master_file = st.file_uploader("📂 Sube tu Catálogo Maestro (Excel .xlsx)", type=["xlsx"], key="master_upload")
    
    if master_file is not None:
        try:
            df_master = pd.read_excel(master_file)
            st.success("¡Archivo maestro cargado con éxito!")
            st.write("Vista previa de tus datos:", df_master.head())
            
            st.markdown("### Selecciona las columnas correspondientes")
            cols = df_master.columns.tolist()
            col_name = st.selectbox("Columna con el Nombre / Descripción del Producto", cols)
            col_code = st.selectbox("Columna con el Código de Barras Oficial", cols)
            
            if st.button("💾 Guardar Catálogo en Memoria"):
                count = 0
                temp_dict = {}
                for _, row in df_master.iterrows():
                    p_name = str(row[col_name]).strip().upper()
                    p_code = clean_barcode(row[col_code])
                    if p_name and p_code != "S/C (Sin Código)":
                        temp_dict[p_name] = p_code
                        count += 1
                
                st.session_state["master_catalog"] = temp_dict
                st.success(f"¡Se han importado y guardado correctamente {count} productos en el catálogo maestro!")
        except Exception as e:
            st.error(f"Error al leer el archivo Excel: {e}")

    if st.session_state["master_catalog"]:
        st.markdown("---")
        st.markdown(f"### 📋 Productos activos en el Catálogo Maestro ({len(st.session_state['master_catalog'])})")
        df_current_master = pd.DataFrame([{"Descripción": k, "Código de Barras": v} for k, v in st.session_state["master_catalog"].items()])
        st.dataframe(df_current_master, use_container_width=True, hide_index=True, height=400)
    else:
        st.info("ℹ️ Aún no hay productos cargados en el catálogo maestro temporal. Sube un archivo Excel para comenzar.")

# ==========================================
# MÓDULO 4: MEMORIA
# ==========================================
elif modulo == "📋 Ver Códigos Almacenados":
    st.markdown("<h2>📋 Memoria de <span style='color: #0284c7;'>Códigos Almacenados</span></h2>", unsafe_allow_html=True)
    b_mem = st.session_state["barcode_memory"]
    if b_mem:
        st.info(f"📊 Total de códigos oficiales almacenados en memoria: **{len(b_mem)}**")
        df_codes = pd.DataFrame([{"Código de Barras Oficial": str(code), "Nombre del Producto": name} for name, code in b_mem.items()])
        st.dataframe(df_codes, use_container_width=True, hide_index=True, height=500)
    else:
        st.warning("⚠️ La memoria está vacía.")
