import io
import json
import os
import time
import re
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

# Diccionario maestro blindado para Álvarez & Sánchez
MASTER_ALVAREZ_SANCHEZ = {
    "VINO TINTO RESERVA CUNE (D.O.RIOJA)": "8410591003045",
    "VINO TINTO MERLOT VIÑA TARAPACA": "7804340909534",
    "VINO TINTO RESERVA CAB SAUV TARAPACA": "7804340909039",
    "VINO TINTO RESERVA CARMENERE TARAPACA": "7804340909010",
    "VINO TINTO RESERVA MERLOT TARAPACA": "78043409041635",
    "VINO TINTO RED BLEND JUAN GIL(JUMILLA)": "8437010482341",
    "VINO TTO ET. AMARILLA JUAN GIL (JUMILLA)": "8437010482297",
    "VINO TTO AZUL JUAN GIL (JUMILLA)": "8437010482273",
    "VINO TTO PLATA JUAN GIL (JUMILLA)": "8437010482280",
    "VINO TINTO MERLOT CALIFORNIA JOSH": "85000020709",
    "VINO TTO CAB SAUV BOURBON RESERV JOSH": "85000020747",
    "VINO TTO CAB SAUV RESERV JOSH": "85000020754",
    "VINO TINTO PINOT NOIR 689 CELLARS": "85000001942",
    "VINO TINTO SIX EIGHT NINE": "051497322618",
    "WHISKY ESCOSÉS MALTA 12 AÑOS GLEN GRANT": "8000040630269",
    "VODKA INFUSIONS CITRUS SKYY": "72105927504",
    "VODKA INFUSIONS RASPBERRY SKYY": "72105920203",
    "VODKA SKYY": "721059007504"
}

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
    digits = re.sub(r'\D', '', s_val)
    if not digits:
        return "S/C (Sin Código)"
    return digits

def run_visual_countdown(total_seconds=30, text_msg="Límite de la capa gratuita alcanzado (15 peticiones por minuto). Pausa de seguridad activa:"):
    st.warning(f"⚠️ {text_msg}")
    bar = st.progress(0)
    status_text = st.empty()
    
    for i in range(total_seconds):
        rem = total_seconds - i
        pct = float((i + 1) / total_seconds)
        bar.progress(pct)
        status_text.text(f"⏳ Faltan {rem} segundos para reintentar automáticamente...")
        time.sleep(1)
        
    bar.empty()
    status_text.empty()

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>⚡ WilPOS</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Automatizador Inteligente (Flash Free)</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Menú de Navegación",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

def parse_empaque_exact(desc, unidad_txt):
    d = str(desc).upper()
    u = str(unidad_txt).upper()
    if "BOT" in u and "VODKA" in d:
        return 1
    if "GLEN GRANT" in d or "WHISKY" in d:
        return 12
    if "CAJA" in u:
        return 12 if "JOSH" in d or "CUNE" in d or "TARAPACA" in d else 6
    return 12

def process_invoice_gemini_flash(file_obj, file_type):
    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini (Configura GEMINI_API_KEY en st.secrets)"

    prompt_text = (
        "Analiza esta factura con precisión milimétrica. "
        "Extrae cada renglón de la tabla con: "
        "1. 'descripcion': texto exacto de la columna 'DESCRIPCION'. "
        "2. 'cantidad': número de la columna 'CANTIDAD'. "
        "3. 'unidad': 'CAJA' o 'BOT.'. "
        "4. 'tamano': texto exacto de la columna 'TAMAÑO'. "
        "5. 'precio_lista': número exacto de la columna 'PRECIO'. "
        "6. 'descuento_porcentaje': porcentaje exacto de la columna 'COM.'. "
        "Devuelve un JSON puro bajo la clave 'items': "
        '{"items": [{"descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            model = genai.GenerativeModel('gemini-3.6-flash')
            
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            
            response = model.generate_content([image_input, prompt_text])
            
            if not response or not response.text:
                raise ValueError("La API de Gemini devolvió una respuesta vacía.")
                
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            
            time.sleep(4.0)
            return json.loads(raw_text.strip()), "✅ Éxito (Gemini Flash Gratuito)"
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "quota" in err_str.lower():
                if intento < 1:
                    run_visual_countdown(30, "Límite de la capa gratuita alcanzado (15 peticiones por minuto). Conteo regresivo visual:")
                    continue
                return None, "QUOTA_EXCEEDED"
            return None, f"Error técnico API: {err_str}"
    return None, "Error desconocido en API"

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.markdown("<h2>📊 Automatizador de Facturas <span style='color: #0284c7;'>(Individual - Flash Gratuito)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesamiento optimizado con Gemini 3.6 Flash y control de tasa integrado.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    c_col1, _ = st.columns([1, 3])
    with c_col1:
        margen_ganancia = st.number_input("⚙️ Ganancia (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    
    uploaded_file = st.file_uploader("📂 Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        if st.button("🚀 Procesar con Gemini Flash (Gratis)"):
            file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
            with st.spinner("Procesando con Gemini Flash..."):
                parsed_data, success_msg = process_invoice_gemini_flash(uploaded_file, file_type)

            if success_msg == "QUOTA_EXCEEDED":
                st.error("⚠️ **Límite de la capa gratuita alcanzado (15 peticiones por minuto).** El contador automático finalizó; por favor presiona nuevamente el botón de procesar.")
            elif not parsed_data or not isinstance(parsed_data, dict):
                st.error(f"⚠️ {success_msg}")
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
                        continue
                    
                    desc = str(item.get("descripcion") or "").strip()
                    if not desc:
                        continue

                    precio_lista = safe_float(item.get("precio_lista") or 0)
                    if precio_lista <= 0:
                        continue

                    upper_desc = desc.upper()
                    extracted_code = "S/C (Sin Código)"
                    for m_key, m_code in MASTER_ALVAREZ_SANCHEZ.items():
                        if m_key in upper_desc or upper_desc in m_key or any(w in upper_desc for w in m_key.split() if len(w) > 5 and w in upper_desc):
                            extracted_code = m_code
                            break

                    extracted_code = clean_barcode(extracted_code)
                    desc_pct = safe_float(item.get("descuento_porcentaje") or 0.0)
                    cant_comprada = safe_int(item.get("cantidad") or 1, 1)
                    unidad_txt = str(item.get("unidad") or "CAJA")
                    tamano_txt = str(item.get("tamano") or "12/75 CL.")
                    
                    empaque_val = parse_empaque_exact(desc, unidad_txt)
                    
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
                        "Código Oficial POS": str(extracted_code),
                        "Nombre Maestro / Artículo": desc,
                        "Cant. Compra": cant_comprada,
                        "Unidad": unidad_txt,
                        "Tamaño/Empaque": tamano_txt,
                        "Empaque Num": empaque_val,
                        "Stock Total": total_unidades_linea,
                        "Desc. %": f"{desc_pct}%",
                        "Costo Unitario": costo,
                        "Precio Venta": precio_venta,
                        "Estado": "OK"
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

                st.info(f"📋 **Auditoría de Lectura:** Se detectaron **{len(data_items)} ítems** en la factura. Procesados y listos: **{len(rows_preview)}**")

                if rows_preview:
                    st.markdown("### ✅ Artículos Procesados Exitosamente")
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
                        ws_prod.cell(row=ws_prod.max_row, column=2).value = str(item_dict["Código Oficial POS"])
                    
                    output = io.BytesIO()
                    wb.save(output)
                    st.download_button("📥 Descargar Excel WilPOS Oficial", output.getvalue(), "Inventario_WilPOS_Actualizado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.markdown("<h2>📂 Procesador por <span style='color: #0284c7;'>Lotes (Flash Free)</span></h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Procesa múltiples facturas aplicando barra de progreso y conteo regresivo visual.</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    l_col1, _ = st.columns([1, 3])
    with l_col1:
        margen_ganancia_lote = st.number_input("⚙️ Ganancia (%) Lote", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
    uploaded_files = st.file_uploader("📂 Sube tu factura (Selección múltiple)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")
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
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["is_live_processing"] = False

        processed_so_far = st.session_state["batch_processed_count"]

        b_col1, b_col2 = st.columns(2)
        iniciar_btn = b_col1.button("🚀 Iniciar Lote Seguro (Gratis)", type="primary")
        reiniciar_lote = b_col2.button("🔄 Reiniciar Lote")

        if reiniciar_lote:
            st.session_state["batch_accumulated_items"] = []
            st.session_state["batch_processed_count"] = 0
            st.session_state["batch_ok_count"] = 0
            st.session_state["is_live_processing"] = False
            if "cached_uploaded_files" in st.session_state:
                del st.session_state["cached_uploaded_files"]
            st.rerun()

        if iniciar_btn:
            st.session_state["is_live_processing"] = True
            st.rerun()

        if st.session_state.get("is_live_processing", False):
            if processed_so_far < total_files:
                file_info = cached_files[processed_so_far]
                st.info(f"⚡ Procesando archivo {processed_so_far + 1} de {total_files}: `{file_info['name']}`...")
                st.progress(processed_so_far / total_files)

                file_bytes_io = io.BytesIO(file_info["bytes"])
                parsed_data, err_msg = process_invoice_gemini_flash(file_bytes_io, file_info["type"])

                if err_msg == "QUOTA_EXCEEDED":
                    run_visual_countdown(30, "Límite temporal alcanzado en lote. Conteo regresivo visual para reanudar:")
                    st.rerun()
                elif "Error técnico" in err_msg:
                    st.error(f"⚠️ {err_msg}")
                    st.session_state["is_live_processing"] = False
                    st.rerun()

                if parsed_data and isinstance(parsed_data, dict):
                    st.session_state["batch_ok_count"] += 1
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
            st.markdown("## 📊 Consolidado de Inventario y Totales")

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

                upper_desc = desc.upper()
                extracted_code = "S/C (Sin Código)"
                for m_key, m_code in MASTER_ALVAREZ_SANCHEZ.items():
                    if m_key in upper_desc or upper_desc in m_key or any(w in upper_desc for w in m_key.split() if len(w) > 5 and w in upper_desc):
                        extracted_code = m_code
                        break

                extracted_code = clean_barcode(extracted_code)
                precio_lista = safe_float(item.get("precio_lista") or 0)
                desc_pct = safe_float(item.get("descuento_porcentaje") or 0.0)
                cant_comprada = safe_int(item.get("cantidad") or 1, 1)
                unidad_txt = str(item.get("unidad") or "CAJA")
                tamano_txt = str(item.get("tamano") or "12/75 CL.")

                empaque_val = parse_empaque_exact(desc, unidad_txt)
                
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
                    "Código Barra": str(extracted_code),
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
                kpi4.metric("💰 Inversión Neta", f"RD$ {inversion_total_lote:,.2f}")

                st.markdown("---")
                st.dataframe(df_final_preview, use_container_width=True, hide_index=True)

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
                    ws_prod.cell(row=ws_prod.max_row, column=2).value = str(row["Código Barra"])

                output = io.BytesIO()
                wb.save(output)
                st.download_button("📥 Descargar Excel Consolidado Final", output.getvalue(), "Inventario_WilPOS_Consolidado_Flash.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
