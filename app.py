import io
import json
import os
import time
from datetime import datetime
import re
import google.generativeai as genai
from PIL import Image
import streamlit as st
import openpyxl
import pandas as pd

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="WilPOS - Facturas Multi-Proveedor", 
    page_icon="⚡", 
    layout="wide"
)

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    .card-container { background-color: #ffffff; border: 1px solid #e2e8f0; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); margin-bottom: 20px; }
    .stButton>button { background: #0284c7; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.5rem; font-weight: 600; }
    .stButton>button:hover { background: #0369a1; color: white; }
    .stDownloadButton>button { background: #10b981; color: white; border: none; border-radius: 8px; padding: 0.5rem 1.2rem; font-weight: 600; }
    </style>
""", unsafe_allow_html=True)

gemini_key = (
    st.secrets.get("GEMINI_API_KEY_PAID") or 
    st.secrets.get("GEMINI_API_KEY") or 
    os.environ.get("GEMINI_API_KEY_PAID") or 
    os.environ.get("GEMINI_API_KEY")
)

if gemini_key:
    genai.configure(api_key=gemini_key)

SUPPLIER_MEMORY_FILE = "proveedores_formatos_memoria.json"
MASTER_CATALOG_FILE = "catalogo_maestro_sistema.json"

def load_json_file(filepath, default_type="dict"):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if default_type == "list" and isinstance(data, list): return data
                if default_type == "dict" and isinstance(data, dict): return data
        except Exception:
            pass
    return [] if default_type == "list" else {}

def save_json_file(filepath, data):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

# ==========================================
# GESTIÓN DE PERFILES Y CATÁLOGO
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "CND / BEES": {
            "nombre": "CND / BEES",
            "tipo_formato": "tique_doble_linea",
            "instruccion_prompt": "Analiza este tique de CND / BEES donde cada ítem tiene dos líneas: la línea 1 con código, unidad (PC o UN) y descripción, y la línea 2 con cantidad, precio unitario (P.Unit) e impuesto neto. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_desglose",
            "instruccion_prompt": "Analiza esta factura de ALVAREZ & SANCHEZ renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'monto_neto'."
        }
    }
    if not loaded_supps:
        loaded_supps = default_profiles
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    loaded_master = load_json_file(MASTER_CATALOG_FILE, "dict")
    base_defaults = {
        "PTE. LIGHT HU 22OZ": "70601561", "PRESIDENTE LIGHT HU 22OZ": "70601561", "PTE. CJ 22OZ": "70601561",
        "PTE. HU 12OZ": "74621774", "PTE. LIGHT HU 12OZ": "74621774", "BRAHMA LIGHT HU 12OZ": "7468973200194",
        "BRAHMA LIGHT HU 16/650M": "7468973200200", "BRAHMA LIGHT 650ML": "7468973200200",
        "CORONA EXTRA 330ML": "7503034941200", "CORONA CERO 355ML": "750304423180", "MICHELOB ULTRA 355ML": "7422110104967",
        "THE ONE HU 12OZ": "74601325", "THE ONE HU 22OZ": "74601127", "CLAMATO COCTEL TOMATE C": "01484035",
        "ENRIQUILLO SODA 400 ML": "7463172803733", "GATORADE FRUIT PUNCH": "7460548000154", "GATORADE NARANJA": "92735",
        "GATORADE UVA": "052000324822", "FOUR LOKO MARACUYA": "849806004962", "FOUR LOKO PONCHE DE FRUTAS": "849806001220",
        "FOUR LOKO GREEN": "849806001855", "FOUR LOKO PURPLE": "849806002746", "FOUR LOKO GOLD": "849806001756",
        "FOUR LOKO SANDIA": "849806001206", "FOUR LOKO WHITE": "849806005754", "ALOE PURE PLUS ORIGINAL": "8809125063011",
        "MY COCO PURE PLUS": "8809125063011",
        # Álvarez & Sánchez
        "SANTA HELENA MERLOT 75 CL": "7804300120986",
        "SANTA HELENA RESERVADO RED BLEND 75 CL": "7804300150082",
        "SANTA HELENA SAUVIGNON BLANC 75 CL": "7804300150041",
        "SANTA HELENA VINO DULCE TINTO 75 CL": "7804300149307",
        "SANTIAGO RUIZ ALBARIÑO 1.5 LT": "8420976010063",
        "SANTIAGO RUIZ ALBARIÑO 375 CL": "8420976010087",
        "SANTIAGO RUIZ ALBARIÑO 75 CL": "842097601070",
        "SCHWEPPES AGUA TONICA 4 PACK 18 CL": "2000011980849",
        "SCHWEPPES TONICA 1 LT": "2117974",
        "SCHWEPPES TONICA ZERO 1 LT": "2138531",
        "SELA BODEGAS RODA VINO TINTO 75 CL": "8014396003073",
        "SOLAN DE CABRAS AGUA MINERAL NAT 1.5 LT": "8436538810767"
    }
    for k, v in base_defaults.items():
        if k not in loaded_master: loaded_master[k] = v
    save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

if "web_encontrado_temporal" not in st.session_state:
    st.session_state["web_encontrado_temporal"] = {}

def safe_float(val, default=0.0):
    try:
        if val is None: return default
        s_val = str(val).replace('$', '').replace(',', '').strip()
        return float(s_val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): 
    return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    s_val = re.sub(r'\D', '', s_val)
    if 7 <= len(s_val) <= 14: return str(s_val)
    return "S/C"

def limpiar_nombre_producto(descripcion_raw, tamano_raw):
    desc = str(descripcion_raw or "").upper().strip()
    desc = re.sub(r'\s+\d{4,6}$', '', desc)
    desc = re.sub(r'\b(PC|UN|CAJA|CAJ|BOT|PZA)\b', '', desc)
    desc = re.sub(r'\s+', ' ', desc).strip()
    
    tam = str(tamano_raw or "").upper().strip()
    m_medida = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z))', tam + " " + desc)
    medida_limpia = m_medida.group(1) if m_medida else ""
    
    if medida_limpia and medida_limpia not in desc:
        return f"{desc} {medida_limpia}".strip()
    
    return desc

def parse_empaque(tamano_txt="", unidad_txt="", descripcion_txt=""):
    unidad_upper = str(unidad_txt or "").upper().strip()
    
    # REGLA ABSOLUTA: Si el tique indica explícitamente que es unidad suelta ("UN"), el empaque es 1
    if unidad_upper == "UN":
        return 1

    combined = f"{str(tamano_txt or '')} {str(unidad_txt or '')} {str(descripcion_txt or '')}".upper()
    
    # Excepción Corona Cero / 4x6 (significa 4 paquetes de 6 = 24 unidades totales)
    if "4X6" in combined:
        return 24

    # 1. Detección por patrones generales con barra
    m_slash = re.search(r'\b(48|24|20|18|16|12|6|4)\s*/', combined)
    if m_slash:
        val = int(m_slash.group(1))
        if val > 1: return val

    if "LP 4" in combined or "4X" in combined: return 24

    # 2. Diccionario de respaldo por palabras clave (solo para cajas / packs PC)
    if "ALOE PURE PLUS" in combined or "MY COCO PURE PLUS" in combined:
        return 20
    if "GATORADE" in combined:
        return 24
    if "FOUR LOKO" in combined:
        return 6
    if "CLAMATO" in combined:
        return 12
    if "ENRIQUILLO" in combined:
        return 24

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_upper = str(nombre_producto or "").upper().strip()
    p_upper = str(presentacion or "").upper().strip()
    combined_query = f"{n_upper} {p_upper}".strip()

    manual_dict = st.session_state.get("codigos_manuales_sesion", {})
    if combined_query in manual_dict: return manual_dict[combined_query]
    if n_upper in manual_dict: return manual_dict[n_upper]

    master_dict = st.session_state.get("master_catalog", {})

    cnd_sinonimos = {
        "PTE. LIGHT HU 22OZ": "70601561", "PRESIDENTE LIGHT HU 22OZ": "70601561", "PTE. CJ 22OZ": "70601561",
        "PTE. HU 12OZ": "74621774", "PTE. LIGHT HU 12OZ": "74621774", "BRAHMA LIGHT HU 12OZ": "7468973200194",
        "BRAHMA LIGHT HU": "7468973200194", "BRAHMA LIGHT HU 16/650M": "7468973200200", "BRAHMA LIGHT 650ML": "7468973200200",
        "CORONA EXTRA 330ML": "7503034941200", "CORONA CERO 355ML": "750304423180", "MICHELOB ULTRA 355ML": "7422110104967",
        "THE ONE HU 12OZ": "74601325", "THE ONE HU 22OZ": "74601127", "CLAMATO COCTEL TOMATE C": "01484035",
        "ENRIQUILLO SODA 400 ML": "7463172803733", "GATORADE FRUIT PUNCH": "7460548000154", "GATORADE NARANJA": "92735",
        "GATORADE UVA": "052000324822", "FOUR LOKO MARACUYA": "849806004962", "FOUR LOKO PONCHE DE FRUTAS": "849806001220",
        "FOUR LOKO GREEN": "849806001855", "FOUR LOKO PURPLE": "849806002746", "FOUR LOKO GOLD": "849806001756",
        "FOUR LOKO SANDIA": "849806001206", "FOUR LOKO WHITE": "849806005754", "ALOE PURE PLUS ORIGINAL": "8809125063011",
        "MY COCO PURE PLUS": "8809125063011",
        # Álvarez & Sánchez
        "SANTA HELENA MERLOT 75 CL": "7804300120986",
        "SANTA HELENA RESERVADO RED BLEND 75 CL": "7804300150082",
        "SANTA HELENA SAUVIGNON BLANC 75 CL": "7804300150041",
        "SANTA HELENA VINO DULCE TINTO 75 CL": "7804300149307",
        "SANTIAGO RUIZ ALBARIÑO 1.5 LT": "8420976010063",
        "SANTIAGO RUIZ ALBARIÑO 375 CL": "8420976010087",
        "SANTIAGO RUIZ ALBARIÑO 75 CL": "842097601070",
        "SCHWEPPES AGUA TONICA 4 PACK 18 CL": "2000011980849",
        "SCHWEPPES TONICA 1 LT": "2117974",
        "SCHWEPPES TONICA ZERO 1 LT": "2138531",
        "SELA BODEGAS RODA VINO TINTO 75 CL": "8014396003073",
        "SOLAN DE CABRAS AGUA MINERAL NAT 1.5 LT": "8436538810767"
    }

    for key, code in cnd_sinonimos.items():
        if key in n_upper or key in combined_query:
            return clean_ean_code(code)

    if master_dict:
        if combined_query in master_dict:
            return clean_ean_code(master_dict[combined_query])
        if n_upper in master_dict:
            return clean_ean_code(master_dict[n_upper])
        for m_name, m_code in master_dict.items():
            if m_name in n_upper or n_upper in m_name:
                return clean_ean_code(m_code)

    return "S/C"

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS System</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Corona Cero y Enriquillo Calibrados</p>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: PROCESAR FACTURA
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador Inteligente con Detección Automática</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura o tique. El sistema calcula costos unitarios por pieza sin ITBIS aplicando el empaque correcto.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""
    if "paginacion_detectada" not in st.session_state: st.session_state["paginacion_detectada"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        st.info("💡 Sube tu documento. El sistema reconocerá el proveedor y validará los códigos EAN automáticamente.")
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu factura o tique (PDF multi-página o Imagen)", type=["pdf", "png", "jpg", "jpeg"])
    
    if archivo_subido is not None:
        if st.button("🚀 Detectar Proveedor y Procesar Documento"):
            with st.spinner("🔍 Analizando documento y extrayendo ítems..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave de API configurada.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in f_type.lower() else Image.open(io.BytesIO(file_bytes))

                    prompt_deteccion = (
                        "Analiza este documento comercial (factura o tique) e identifica estrictamente el nombre comercial del proveedor emisor. "
                        "Devuelve únicamente un JSON con esta estructura: {\"proveedor_detectado\": \"NOMBRE DEL PROVEEDOR\"}"
                    )
                    
                    response_det = model.generate_content([image_input, prompt_deteccion])
                    raw_det_text = response_det.text.strip()
                    if raw_det_text.startswith("```json"): raw_det_text = raw_det_text[7:]
                    if raw_det_text.endswith("```"): raw_det_text = raw_det_text[:-3]
                    
                    det_json = json.loads(raw_det_text.strip())
                    nombre_detectado_raw = str(det_json.get("proveedor_detectado", "PROVEEDOR GENERAL")).upper().strip()

                    supp_mem = st.session_state["supplier_memory"]
                    prov_encontrado = None
                    
                    for p_key in supp_mem.keys():
                        if p_key in nombre_detectado_raw or nombre_detectado_raw in p_key:
                            prov_encontrado = p_key
                            break
                    
                    if not prov_encontrado:
                        prov_encontrado = nombre_detectado_raw
                        supp_mem[prov_encontrado] = {
                            "nombre": prov_encontrado,
                            "tipo_formato": "factura_desglose",
                            "instruccion_prompt": f"Analiza esta factura de {prov_encontrado} renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
                        }
                        st.session_state["supplier_memory"] = supp_mem
                        save_json_file(SUPPLIER_MEMORY_FILE, supp_mem)

                    instruccion_proveedor = supp_mem[prov_encontrado].get("instruccion_prompt", "Extrae todos los ítems.")

                    prompt_unificado = (
                        f"Estás procesando un tique o factura del proveedor: '{prov_encontrado}'. "
                        f"Instrucción específica: {instruccion_proveedor} "
                        "IMPORTANTE (Estructura de tique CND/BEES): Cada producto tiene una línea de texto arriba (con el nombre y formato ej. 24/591) y una línea abajo con la cantidad, la unidad (PC o UN), el precio unitario exacto (P.Unit) y el importe neto. "
                        "Extrae 'precio_unitario' exactamente como aparece en la columna P.Unit del tique y la 'unidad' (PC o UN). "
                        "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                        '{"paginacion": "1 de 1", "proveedor_detectado": "' + prov_encontrado + '", "subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"descripcion": "...", "tamano": "...", "cantidad": 1.0, "unidad": "PC", "precio_unitario": 0.0, "monto_neto": 0.0}]}. '
                        "Respuesta JSON pura."
                    )

                    archivo_subido.seek(0)
                    response = model.generate_content([image_input, prompt_unificado])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("```"): raw_text = raw_text[:-3]
                    
                    parsed_json = json.loads(raw_text.strip())

                    st.session_state["factura_data"] = parsed_json
                    st.session_state["prov_activo"] = prov_encontrado
                    st.session_state["paginacion_detectada"] = str(parsed_json.get("paginacion", "1 de 1"))
                    
                    st.success(f"🎯 **¡Proveedor Detectado!** Perfil aplicado: **{prov_encontrado}** ({len(parsed_json.get('items', []))} renglones).")
                except Exception as e:
                    st.error(f"⚠️ Error al procesar: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state["factura_data"] is not None:
        data_resp = st.session_state["factura_data"]
        items = data_resp.get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")
        pag_info = str(st.session_state.get("paginacion_detectada", "1 de 1"))
        
        subtotal_val = safe_float(data_resp.get("subtotal"))
        itbis_val = safe_float(data_resp.get("itbis"))
        total_descuentos = safe_float(data_resp.get("descuentos"))
        total_val = safe_float(data_resp.get("total"))

        total_importe_neto = sum(safe_float(i.get("monto_neto") or i.get("impuesto_neto")) for i in items)
        if subtotal_val == 0.0 and items: subtotal_val = total_importe_neto
        if total_val == 0.0 and items: total_val = subtotal_val * 1.18

        st.markdown(f"### 📊 Dashboard Financiero | Proveedor: {prov_actual} (Pág. {pag_info})")
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1: st.metric(label="Subtotal / Bruto", value=f"${subtotal_val:,.2f}")
        with col_m2: st.metric(label="ITBIS Total", value=f"${itbis_val:,.2f}")
        with col_m3: st.metric(label="Descuentos", value=f"${total_descuentos:,.2f}")
        with col_m4: st.metric(label="Total General", value=f"${total_val:,.2f}")
            
        st.markdown("---")

        if items:
            st.markdown(f"### 📋 Detalle de Renglones Extraídos ({len(items)} ítems)")
            
            preview_rows = []
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            master_dict = st.session_state.get("master_catalog", {})
            nombres_maestro_lista = list(master_dict.keys())
            web_temp = st.session_state.get("web_encontrado_temporal", {})

            for idx, item in enumerate(items, start=1):
                raw_desc = item.get("descripcion", "")
                raw_tam = item.get("tamano", "")
                
                nombre_completo = limpiar_nombre_producto(raw_desc, raw_tam)
                m_med = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z))', str((raw_tam or "") + " " + (raw_desc or "")).upper())
                presentacion_limpia = m_med.group(1) if m_med else (raw_tam if raw_tam else "S/P")

                codigo_final = buscar_en_catalogo_maestro(nombre_completo, presentacion_limpia)

                cant_compra = safe_float(item.get("cantidad"), 1.0)
                unidad = str(item.get("unidad", ""))
                empaque = parse_empaque(raw_tam, unidad, raw_desc)
                total_unidades = int(cant_compra * empaque)

                p_unit_extraido = safe_float(item.get("precio_unitario"), 0.0)
                monto_neto_linea = safe_float(
                    item.get("monto_neto") or 
                    item.get("impuesto_neto") or 
                    item.get("valor"), 
                    0.0
                )

                if p_unit_extraido > 0:
                    costo_unitario_real = round(p_unit_extraido / empaque, 2) if empaque > 1 else p_unit_extraido
                elif monto_neto_linea > 0 and total_unidades > 0:
                    costo_unitario_real = round(monto_neto_linea / total_unidades, 2)
                else:
                    costo_unitario_real = 0.0

                if costo_unitario_real > 0:
                    precio_con_utilidad = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(precio_con_utilidad * 1.18)
                else:
                    precio_venta = 0.0

                display_codigo = codigo_final
                if codigo_final == "S/C":
                    st.markdown("---")
                    
                    key_temp = f"{idx}_{nombre_completo}"
                    
                    col_c1, col_c2, col_c3 = st.columns([2, 1, 1])
                    with col_c1:
                        st.markdown(f"⚠️ **{nombre_completo} ({presentacion_limpia})** sin código.")
                        sel_maestro = st.selectbox(
                            f"Seleccionar del Catálogo Maestro",
                            ["-- Buscar en Maestro --"] + nombres_maestro_lista,
                            key=f"sel_maestro_{idx}_{nombre_completo}"
                        )
                        if sel_maestro != "-- Buscar en Maestro --":
                            codigo_seleccionado = master_dict[sel_maestro]
                            display_codigo = codigo_seleccionado
                            st.session_state["codigos_manuales_sesion"][nombre_completo] = display_codigo
                            master_dict[nombre_completo] = display_codigo
                            save_json_file(MASTER_CATALOG_FILE, master_dict)
                            st.success(f"¡Relacionado con '{sel_maestro}' y guardado en el Catálogo Maestro!")
                            st.rerun()
                    with col_c2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if key_temp not in web_temp:
                            if st.button(f"🌐 Buscar en Web", key=f"web_btn_{idx}_{nombre_completo}"):
                                with st.spinner(f"Consultando bases de datos de códigos UPC/EAN..."):
                                    try:
                                        model_web = genai.GenerativeModel('gemini-3.8-flash')
                                        prompt_web = (
                                            f"Actúa como un experto en logística de inventarios y códigos de barras de productos de consumo masivo (supermercados y POS). "
                                            f"Busca en internet el código de barras UPC o EAN oficial exacto para la unidad individual del producto: '{nombre_completo} con presentación {presentacion_limpia}'. "
                                            "IMPORTANTE: Asegúrate de que corresponda al código de barras de la unidad/botella y no a una caja de empaque múltiple. "
                                            "Devuelve estrictamente y únicamente el número de código de barras puro (de 8 a 14 dígitos). No agregues texto ni explicaciones."
                                        )
                                        res_web = model_web.generate_content(prompt_web)
                                        codigo_web = clean_ean_code(res_web.text.strip())
                                        if codigo_web != "S/C":
                                            web_temp[key_temp] = codigo_web
                                            st.session_state["web_encontrado_temporal"] = web_temp
                                            st.rerun()
                                        else:
                                            st.warning("No se halló el código en la web automáticamente.")
                                    except Exception:
                                        st.error("Error al consultar la web.")
                        else:
                            codigo_hallado = web_temp[key_temp]
                            st.info(f"✨ Hallado: **{codigo_hallado}**")
                            
                            sub_col_b1, sub_col_b2 = st.columns(2)
                            with sub_col_b1:
                                if st.button(f"✅ Confirmar", key=f"conf_btn_{idx}_{nombre_completo}"):
                                    display_codigo = codigo_hallado
                                    st.session_state["codigos_manuales_sesion"][nombre_completo] = display_codigo
                                    master_dict[nombre_completo] = display_codigo
                                    save_json_file(MASTER_CATALOG_FILE, master_dict)
                                    del web_temp[key_temp]
                                    st.session_state["web_encontrado_temporal"] = web_temp
                                    st.success(f"¡Agregado y asignado!")
                                    st.rerun()
                            with sub_col_b2:
                                if st.button(f"❌ Rechazar", key=f"rej_btn_{idx}_{nombre_completo}"):
                                    del web_temp[key_temp]
                                    st.session_state["web_encontrado_temporal"] = web_temp
                                    st.warning("Resultado rechazado. Puedes volver a buscar.")
                                    st.rerun()
                    with col_c3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        codigo_manual_input = st.text_input("O ingresa código manual", key=f"manual_input_{idx}_{nombre_completo}", placeholder="Ej. 052000324822")
                        if codigo_manual_input and len(codigo_manual_input.strip()) >= 7:
                            clean_m = clean_ean_code(codigo_manual_input)
                            if clean_m != "S/C":
                                display_codigo = clean_m
                                st.session_state["codigos_manuales_sesion"][nombre_completo] = display_codigo
                                master_dict[nombre_completo] = display_codigo
                                save_json_file(MASTER_CATALOG_FILE, master_dict)
                                st.rerun()

                preview_rows.append({
                    "No.": idx, "Producto": nombre_completo, "Código EAN Asignado": display_codigo,
                    "Cant. Compra": cant_compra, "Empaque": empaque, "Stock (Unidades)": total_unidades,
                    "Costo Unit. Real": costo_unitario_real, "Precio Venta": precio_venta
                })

                ws.append([
                    nombre_completo, presentacion_limpia, str(display_codigo), prov_actual, "producto",
                    precio_venta, costo_unitario_real, total_unidades, 0.18, "unidad", empaque
                ])

            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            st.download_button(
                label=f"📥 Descargar Excel Importable - {prov_actual} (Pág. {pag_info})",
                data=excel_buffer.getvalue(),
                file_name=f"Inventario_Master_{prov_actual.replace(' ', '_')}_Pag_{pag_info.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==========================================
# MÓDULO 2: CATÁLOGO MAESTRO EAN
# ==========================================
elif menu_opcion == "📁 Catálogo Maestro EAN":
    st.markdown("<h2>📁 Gestión del Catálogo Maestro EAN</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Administra y registra nuevos productos con sus códigos de barra o SAP oficiales.</p>", unsafe_allow_html=True)
    st.markdown("---")

    master_dict = st.session_state.get("master_catalog", {})

    with st.expander("➕ Agregar o Actualizar Producto Manualmente en el Maestro", expanded=True):
        with st.form("form_agregar_maestro"):
            col_m1, col_m2 = st.columns([2, 1])
            with col_m1:
                nuevo_prod_nombre = st.text_input("Nombre / Descripción del Producto (Ej: SCHWEPPES TONICA 1 LT)")
            with col_m2:
                nuevo_prod_codigo = st.text_input("Código EAN / SAP Oficial (Ej: 2117974)")
            
            btn_guardar_maestro = st.form_submit_button("💾 Guardar en Catálogo Maestro")
            if btn_guardar_maestro:
                clean_name = nuevo_prod_nombre.upper().strip()
                clean_code = clean_ean_code(nuevo_prod_codigo)
                if clean_name and clean_code != "S/C":
                    master_dict[clean_name] = clean_code
                    st.session_state["master_catalog"] = master_dict
                    save_json_file(MASTER_CATALOG_FILE, master_dict)
                    st.success(f"✅ ¡Producto **{clean_name}** guardado con éxito con el código **{clean_code}**!")
                    st.rerun()
                else:
                    st.error("⚠️ Por favor ingresa un nombre válido y un código EAN/SAP correcto.")

    if master_dict:
        st.markdown(f"### 📋 Productos Registrados en el Catálogo ({len(master_dict):,} registros)")
        df_show = pd.DataFrame([{"Producto / Descripción": k, "Código EAN/SAP Oficial": v} for k, v in master_dict.items()])
        st.dataframe(df_show, use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 3: GESTIONAR PROVEEDORES
# ==========================================
elif menu_opcion == "🏢 Gestionar Proveedores":
    st.markdown("<h2>🏢 Configuración de Perfiles por Proveedor</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Cada proveedor mantiene su propia regla de extracción intacta y respaldada.</p>", unsafe_allow_html=True)
    st.markdown("---")

    supps = st.session_state["supplier_memory"]
    for p_name, p_data in list(supps.items()):
        with st.expander(f"🏢 Proveedor: {p_name}"):
            with st.form(f"form_prov_{p_name}"):
                nuevo_nombre = st.text_input("Nombre del Proveedor", value=p_data.get("nombre", p_name))
                nueva_instruccion = st.text_area("Instrucción / Prompt de Formato Exclusivo", value=p_data.get("instruccion_prompt", ""), height=120)
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    btn_guardar = st.form_submit_button("💾 Guardar Cambios")
                with col_btn2:
                    btn_eliminar = st.form_submit_button("🗑️ Eliminar Perfil")
                
                if btn_guardar:
                    supps[p_name]["nombre"] = nuevo_nombre
                    supps[p_name]["instruccion_prompt"] = nueva_instruccion
                    if nuevo_nombre != p_name:
                        supps[nuevo_nombre] = supps.pop(p_name)
                    st.session_state["supplier_memory"] = supps
                    save_json_file(SUPPLIER_MEMORY_FILE, supps)
                    st.success(f"✅ ¡Perfil de **{nuevo_nombre}** guardado con éxito!")
                    st.rerun()
                    
                if btn_eliminar:
                    if p_name in supps:
                        supps.pop(p_name)
                        st.session_state["supplier_memory"] = supps
                        save_json_file(SUPPLIER_MEMORY_FILE, supps)
                        st.warning(f"⚠️ Perfil de {p_name} eliminado.")
                        st.rerun()

    with st.expander("➕ Agregar Nuevo Proveedor Manualmente"):
        with st.form("form_nuevo_proveedor_manual"):
            n_prov = st.text_input("Nombre del Proveedor (Ej: CASA BRUGAL)")
            n_inst = st.text_area("Instrucción de Formato para este Proveedor", value="Analiza la factura de este proveedor y extrae descripción, tamaño, cantidad, unidad y monto neto.")
            btn_crear = st.form_submit_button("Crear Perfil de Proveedor")
            if btn_crear:
                clean_p = n_prov.upper().strip()
                if clean_p:
                    supps[clean_p] = {"nombre": clean_p, "tipo_formato": "personalizado", "instruccion_prompt": n_inst}
                    st.session_state["supplier_memory"] = supps
                    save_json_file(SUPPLIER_MEMORY_FILE, supps)
                    st.success(f"✅ ¡Perfil creado para {clean_p}!")
                    st.rerun()
