import io
import json
import os
import time
import unicodedata
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
    page_title="WilPOS - Sistema Multi-Proveedor Maestro", 
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
                if default_type == "dict":
                    if isinstance(data, dict):
                        return {str(k).upper().strip(): v for k, v in data.items() if k}
                    return {}
                elif default_type == "list":
                    if isinstance(data, list):
                        return data
                    return []
        except Exception:
            pass
    return {} if default_type == "dict" else []

def save_json_file(filepath, data):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

# ==========================================
# UTILIDADES Y FUNCIONES AUXILIARES
# ==========================================
def safe_float(val, default=0.0):
    try:
        if val is None: return default
        s_val = str(val).replace('$', '').replace('%', '').replace(',', '').strip()
        return float(s_val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): 
    return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    s_val = re.sub(r'\D', '', s_val)
    if 7 <= len(s_val) <= 14: 
        return str(s_val)
    return "S/C"

def normalizar_texto(texto):
    if not texto: return ""
    t = str(texto).upper().strip()
    t = ''.join(c for c in unicodedata.normalize('NFD', t) if unicodedata.category(c) != 'Mn')
    t = re.sub(r'([A-Z])(\d)', r'\1 \2', t)
    t = re.sub(r'(\d)([A-Z])', r'\1 \2', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def limpiar_nombre_y_extraer_presentacion(proveedor_activo, descripcion_raw, tamano_raw=""):
    t_norm = normalizar_texto(descripcion_raw)
    t_tam = normalizar_texto(tamano_raw)
    prov_up = normalizar_texto(proveedor_activo)
    
    combined_raw = f"{t_norm} {t_tam}"
    
    m_med = re.search(r'(\d+)\s*(?:ML|L|LT|G|KG|OZ|CL)', combined_raw)
    if m_med:
        num_val = m_med.group(1)
        if "L" in combined_raw[m_med.end()-3:m_med.end()] and int(num_val) <= 2:
            presentacion = f"{int(num_val)*1000}ML" if int(num_val) <= 2 else f"{num_val}ML"
        else:
            presentacion = f"{num_val}ML" if "CL" not in combined_raw else f"{int(num_val)*10}ML"
    else:
        presentacion = "750ML" if "EL CATADOR" in prov_up else "UN"

    if "75CL" in combined_raw: presentacion = "750ML"
    if "1L" in combined_raw or "1LT" in combined_raw or "1000ML" in combined_raw: presentacion = "1000ML"

    desc_limpia = t_norm
    desc_limpia = re.sub(r'\b(CAJA\s*\d*|CJ\s*\d*\s*BOT|BOT|BOTELLA|LATA|LATAS|\d+\s*X\s*\d+\s*(?:ML|CL|L)|\d+/\s*\d+\s*(?:ML|CL|L|750ML)|750\s*ML|75\s*CL|750ML|75CL|1000\s*ML|1000ML|1L|1LT|500\s*ML|500ML|375\s*ML|375ML|700\s*ML|700ML|50\s*ML|50ML|1250\s*ML|1250ML|\d+OZ|\d+\s*PZAS|\d+\s*PZA|\d+P\b|\b4P\b|\b12P\b|\b\d+X\d+[A-Z]*\b)\b', '', desc_limpia)
    desc_limpia = re.sub(r'\s+', ' ', desc_limpia).strip()
    
    if presentacion != "UN" and presentacion not in desc_limpia:
        desc_limpia = f"{desc_limpia} {presentacion}".strip()
    else:
        for p_test in ["500ML", "375ML", "50ML", "750ML", "1000ML", "700ML", "1250ML"]:
            if desc_limpia.count(p_test) > 1:
                desc_limpia = desc_limpia.replace(p_test, "", desc_limpia.count(p_test) - 1).strip()

    return desc_limpia, presentacion

def parse_empaque_proveedor(proveedor_nombre, unidad_txt="", tamano_txt="", descripcion_txt=""):
    unidad_norm = normalizar_texto(unidad_txt)
    combined = normalizar_texto(f"{unidad_txt} {tamano_txt} {descripcion_txt}")
    
    m_mult = re.search(r'\b(\d+)\s*[xX]\s*\d+', unidad_norm)
    if m_mult:
        val = int(m_mult.group(1))
        if val > 0: return val

    m_mult_comb = re.search(r'\b(\d+)\s*[xX]\s*\d+', combined)
    if m_mult_comb:
        val = int(m_mult_comb.group(1))
        if val > 0: return val

    if any(m in unidad_norm for m in ["ML", "CL", "L", "OZ"]) and "X" not in unidad_norm:
        return 1

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_norm = normalizar_texto(nombre_producto)
    p_norm = normalizar_texto(presentacion)
    
    master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")
    if not master_dict:
        return "S/C"

    query_completa = f"{n_norm} {p_norm}".strip()
    for m_key, m_code in master_dict.items():
        m_key_up = normalizar_texto(m_key)
        if m_key_up == query_completa:
            return clean_ean_code(m_code)

    return "S/C"

# ==========================================
# GESTIÓN DE PERFILES Y CATÁLOGO
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "UNITED BRANDS S A": {
            "nombre": "UNITED BRANDS S A",
            "tipo_formato": "factura_tabla_united_brands",
            "instruccion_prompt": "Analiza esta página de la factura de UNITED BRANDS S A renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "BEPENSA DOMINICANA SA": {
            "nombre": "BEPENSA DOMINICANA SA",
            "tipo_formato": "factura_tique_bepensa",
            "instruccion_prompt": "Analiza esta página de la factura de BEPENSA DOMINICANA SA renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "EL CATADOR": {
            "nombre": "EL CATADOR",
            "tipo_formato": "factura_cajas_descuento",
            "instruccion_prompt": "Analiza esta página de la factura de EL CATADOR renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "PRICESMART": {
            "nombre": "PRICESMART",
            "tipo_formato": "factura_tique_unidades",
            "instruccion_prompt": "Analiza esta página del comprobante de PRICESMART renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario' y 'monto_neto'."
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_desglose_descuentos",
            "instruccion_prompt": "Analiza esta página de la factura de ALVAREZ & SANCHEZ renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        }
    }
    if not loaded_supps or "UNITED BRANDS S A" not in loaded_supps:
        loaded_supps.update(default_profiles)
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    loaded_master = load_json_file(MASTER_CATALOG_FILE, "dict")
    if not loaded_master:
        loaded_master = {}
        save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

if "paginas_procesadas_historial" not in st.session_state:
    st.session_state["paginas_procesadas_historial"] = set()

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS Multi-Proveedor</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura (Por Página)", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Control Total Activo</p>", unsafe_allow_html=True)

if st.sidebar.button("🔄 Reiniciar Todo el Sistema (Maestro y Sesión)"):
    st.session_state["paginas_procesadas_historial"] = set()
    st.session_state["codigos_manuales_sesion"] = {}
    st.session_state["factura_data"] = None
    save_json_file(MASTER_CATALOG_FILE, {})
    st.success("¡Sistema reiniciado por completo con éxito!")
    time.sleep(0.5)
    st.rerun()

# ==========================================
# MÓDULO 1: PROCESAR FACTURA POR PÁGINA
# ==========================================
if menu_opcion == "📄 Procesar Factura (Por Página)":
    st.markdown("<h2>📄 Procesador de Facturas (Multi-Proveedor)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura. Las ediciones se realizan mediante doble clic de forma temporal en sesión y nunca alteran el maestro.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        st.info("💡 Sube tu factura (PDF multi-página o imagen).")
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu archivo", type=["pdf", "png", "jpg", "jpeg"])
    
    pagina_a_procesar = 1
    if archivo_subido is not None and archivo_subido.name.endswith('.pdf'):
        try:
            import fitz # PyMuPDF
            archivo_subido.seek(0)
            doc_pdf = fitz.open(stream=archivo_subido.read(), filetype="pdf")
            total_paginas = len(doc_pdf)
            st.markdown(f"📄 **PDF detectado con {total_paginas} página(s).**")
            pagina_a_procesar = st.selectbox("Selecciona la página exacta a procesar:", list(range(1, total_paginas + 1)))
        except Exception:
            pagina_a_procesar = st.number_input("Número de página a procesar:", min_value=1, value=1, step=1)

    if archivo_subido is not None:
        if st.button(f"🚀 Procesar Página #{pagina_a_procesar}"):
            with st.spinner(f"🔍 Analizando página #{pagina_a_procesar}..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave de API configurada.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    
                    target_image = None
                    if "pdf" in f_type.lower():
                        import fitz
                        doc_pdf = fitz.open(stream=file_bytes, filetype="pdf")
                        page_obj = doc_pdf[pagina_a_procesar - 1]
                        pix = page_obj.get_pixmap(dpi=150)
                        target_image = Image.open(io.BytesIO(pix.tobytes("jpeg")))
                    else:
                        target_image = Image.open(io.BytesIO(file_bytes))

                    prompt_deteccion = (
                        "Analiza esta página de documento comercial e identifica estrictamente el nombre comercial del proveedor emisor y el número de factura si lo hubiera. "
                        "Devuelve únicamente un JSON con esta estructura: {\"proveedor_detectado\": \"NOMBRE DEL PROVEEDOR\", \"numero_factura\": \"NUMERO\"}"
                    )
                    
                    response_det = model.generate_content([target_image, prompt_deteccion])
                    raw_det_text = response_det.text.strip()
                    if raw_det_text.startswith("```json"): raw_det_text = raw_det_text[7:]
                    if raw_det_text.endswith("```"): raw_det_text = raw_det_text[:-3]
                    
                    det_json = json.loads(raw_det_text.strip())
                    nombre_detectado_raw = str(det_json.get("proveedor_detectado", "PROVEEDOR GENERAL")).upper().strip()
                    num_factura_detectada = str(det_json.get("numero_factura", "S/N")).upper().strip()

                    firma_pagina = f"{archivo_subido.name}_{num_factura_detectada}_PAG_{pagina_a_procesar}"
                    
                    if firma_pagina in st.session_state["paginas_procesadas_historial"]:
                        st.warning(f"⚠️ La página #{pagina_a_procesar} ya fue procesada previamente en esta sesión.")

                    supp_mem = st.session_state["supplier_memory"]
                    if not isinstance(supp_mem, dict): supp_mem = {}
                    
                    prov_encontrado = None
                    for p_key in supp_mem.keys():
                        if p_key in nombre_detectado_raw or nombre_detectado_raw in p_key:
                            prov_encontrado = p_key
                            break
                    
                    if not prov_encontrado:
                        prov_encontrado = nombre_detectado_raw
                        supp_mem[prov_encontrado] = {
                            "nombre": prov_encontrado,
                            "tipo_formato": "factura_desglose_personalizado",
                            "instruccion_prompt": f"Analiza esta página de {prov_encontrado} renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
                        }
                        st.session_state["supplier_memory"] = supp_mem
                        save_json_file(SUPPLIER_MEMORY_FILE, supp_mem)

                    prov_dict_data = supp_mem.get(prov_encontrado, {})
                    if not isinstance(prov_dict_data, dict): prov_dict_data = {}
                    instruccion_proveedor = prov_dict_data.get("instruccion_prompt", "Extrae todos los ítems.")

                    prompt_unificado = (
                        f"Estás procesando la página {pagina_a_procesar} de una factura del proveedor: '{prov_encontrado}'. "
                        f"Instrucción específica de su perfil: {instruccion_proveedor} "
                        "Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto' (Impuesto Neto sin ITBIS y con descuento aplicado). "
                        "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                        '{"paginacion": "' + str(pagina_a_procesar) + '", "proveedor_detectado": "' + prov_encontrado + '", "subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"descripcion": "...", "tamano": "750ML", "cantidad": 1.0, "unidad": "750ML", "precio_unitario": 100.0, "descuento_porcentaje": 0.0, "monto_neto": 100.0}]}. '
                        "Respuesta JSON pura."
                    )

                    response = model.generate_content([target_image, prompt_unificado])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("```"): raw_text = raw_text[:-3]
                    
                    parsed_json = json.loads(raw_text.strip())

                    st.session_state["paginas_procesadas_historial"].add(firma_pagina)
                    st.session_state["factura_data"] = parsed_json
                    st.session_state["prov_activo"] = prov_encontrado
                    
                    st.success(f"🎯 **¡Página #{pagina_a_procesar} procesada con éxito!** Proveedor: **{prov_encontrado}** ({len(parsed_json.get('items', []))} renglones).")
                except Exception as e:
                    st.error(f"⚠️ Error al procesar página: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state["factura_data"] is not None:
        data_resp = st.session_state["factura_data"]
        items = data_resp.get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")
        pag_info = str(data_resp.get("paginacion", "1"))
        
        subtotal_val = safe_float(data_resp.get("subtotal"))
        itbis_val = safe_float(data_resp.get("itbis"))
        total_descuentos = safe_float(data_resp.get("descuentos"))
        total_val = safe_float(data_resp.get("total"))

        total_importe_neto = sum(safe_float(i.get("monto_neto") or (safe_float(i.get("precio_unitario")) * safe_float(i.get("cantidad")))) for i in items)
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
            lista_codigos_pagina = []
            master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")
            nombres_maestro_lista = list(master_dict.keys())

            for idx, item in enumerate(items, start=1):
                raw_desc = str(item.get("descripcion", ""))
                raw_tam = str(item.get("tamano", ""))
                unidad = str(item.get("unidad", ""))
                
                nombre_limpio, presentacion_limpia = limpiar_nombre_y_extraer_presentacion(prov_actual, raw_desc, raw_tam)
                
                cant_compra = safe_float(item.get("cantidad"), 1.0)
                empaque = parse_empaque_proveedor(prov_actual, unidad, raw_tam, raw_desc)
                total_unidades = int(cant_compra * empaque)

                codigo_final = buscar_en_catalogo_maestro(nombre_limpio, presentacion_limpia)

                if nombre_limpio in st.session_state["codigos_manuales_sesion"]:
                    codigo_final = st.session_state["codigos_manuales_sesion"][nombre_limpio]

                monto_neto_linea = safe_float(item.get("monto_neto"), 0.0)
                p_unit_extraido = safe_float(item.get("precio_unitario"), 0.0)
                desc_pct = safe_float(item.get("descuento_porcentaje"), 0.0)

                if monto_neto_linea > 0 and total_unidades > 0:
                    costo_unitario_real = round(monto_neto_linea / total_unidades, 4)
                elif p_unit_extraido > 0:
                    precio_neto_caja = p_unit_extraido * (1 - (desc_pct / 100.0))
                    costo_unitario_real = round(precio_neto_caja / empaque, 4) if empaque > 1 else precio_neto_caja
                else:
                    costo_unitario_real = 0.0

                if costo_unitario_real > 0:
                    precio_con_utilidad = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(precio_con_utilidad * 1.18)
                else:
                    precio_venta = 0.0

                lista_codigos_pagina.append({
                    "idx": idx, "Producto": nombre_limpio, "Presentación": presentacion_limpia,
                    "Código EAN": str(codigo_final), "Cant. Compra": cant_compra, "Empaque": empaque,
                    "Stock Unidades": total_unidades, "Costo Unit. Real": costo_unitario_real, "Precio Venta": precio_venta
                })

            # ==========================================
            # TABLA EDITABLE INTERACTIVA (st.data_editor con Doble Clic)
            # ==========================================
            df_preview_original = pd.DataFrame(lista_codigos_pagina)
            
            df_editado = st.data_editor(
                df_preview_original,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["idx", "Producto", "Presentación", "Cant. Compra", "Empaque", "Stock Unidades", "Costo Unit. Real", "Precio Venta"],
                column_config={
                    "Código EAN": st.column_config.TextColumn(
                        "Código EAN",
                        help="Haz doble clic en la celda para escribir o corregir el código EAN directamente en esta sesión.",
                        max_chars=14,
                        validate="^\\d+$"
                    )
                },
                key="editor_tabla_inventario"
            )

            # Sincronizar códigos editados EXCLUSIVAMENTE EN SESIÓN (sin tocar el archivo maestro)
            for i, row in df_editado.iterrows():
                p_name = row["Producto"]
                nuevo_c_editado = clean_ean_code(row["Código EAN"])
                if nuevo_c_editado != "S/C":
                    st.session_state["codigos_manuales_sesion"][p_name] = nuevo_c_editado

            # ==========================================
            # ASISTENTE PARA ÍTEMS SIN CÓDIGO (S/C)
            # ==========================================
            for row_item in lista_codigos_pagina:
                nombre_limpio = row_item["Producto"]
                presentacion_limpia = row_item["Presentación"]
                actual_c = df_editado.loc[df_editado["Producto"] == nombre_limpio, "Código EAN"].values[0]
                
                if actual_c == "S/C":
                    st.markdown(f"⚠️ **{nombre_limpio} ({presentacion_limpia})** sin código EAN asignado.")
                    col_c1, col_c2, col_c3 = st.columns([2, 1, 1])
                    with col_c1:
                        sel_maestro = st.selectbox("Seleccionar de Catálogo Maestro", ["-- Buscar en Maestro --"] + nombres_maestro_lista, key=f"sel_m_{nombre_limpio}")
                        if sel_maestro != "-- Buscar en Maestro --" and sel_maestro in master_dict:
                            asig_c = master_dict[sel_maestro]
                            st.session_state["codigos_manuales_sesion"][nombre_limpio] = asig_c
                            st.success(f"¡Código asignado para esta factura!")
                            time.sleep(0.3)
                            st.rerun()
                    with col_c2:
                        codigo_manual_input = st.text_input("O ingresar manual", key=f"man_{nombre_limpio}", placeholder="EAN...")
                        if codigo_manual_input and len(codigo_manual_input.strip()) >= 7:
                            clean_m = clean_ean_code(codigo_manual_input)
                            if clean_m != "S/C":
                                st.session_state["codigos_manuales_sesion"][nombre_limpio] = clean_m
                                st.success("¡Guardado para esta factura!")
                                time.sleep(0.3)
                                st.rerun()
                    with col_c3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        q_b = f"EAN barcode {nombre_limpio} {presentacion_limpia}".replace(" ", "+")
                        st.markdown(f"[🌐 Buscar en Google](https://www.google.com/search?q={q_b})", unsafe_allow_html=True)

            # Generación de Excel final con los datos limpios y editados
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo Unitario', 'Stock (Unidades)', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            for idx, row in df_editado.iterrows():
                p_nom = row["Producto"]
                p_pres = row["Presentación"]
                p_ean = row["Código EAN"]
                if p_nom in st.session_state["codigos_manuales_sesion"]:
                    p_ean = st.session_state["codigos_manuales_sesion"][p_nom]

                ws.append([
                    p_nom, p_pres, str(p_ean), prov_actual, "producto",
                    row["Precio Venta"], row["Costo Unit. Real"], row["Stock Unidades"], 0.18, "unidad", row["Empaque"]
                ])

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            st.download_button(
                label=f"📥 Descargar Excel - {prov_actual} (Pág. {pag_info})",
                data=excel_buffer.getvalue(),
                file_name=f"Inventario_Master_{prov_actual.replace(' ', '_')}_Pag_{pag_info}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==========================================
# MÓDULO 2: CATÁLOGO MAESTRO EAN
# ==========================================
elif menu_opcion == "📁 Catálogo Maestro EAN":
    st.markdown("<h2>📁 Gestión, Carga y Limpieza del Archivo Maestro EAN</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Este es el único lugar donde se alimenta y actualiza el archivo maestro oficial de forma manual.</p>", unsafe_allow_html=True)
    st.markdown("---")

    master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")

    col_rst1, col_rst2 = st.columns([3, 1])
    with col_rst2:
        if st.button("⚠️ Vaciar Archivo Maestro"):
            save_json_file(MASTER_CATALOG_FILE, {})
            st.success("¡Catálogo maestro vaciado correctamente!")
            time.sleep(0.5)
            st.rerun()

    # ==========================================
    # CARGA MASIVA DE EXCEL AL MAESTRO
    # ==========================================
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📂 Carga Masiva de Catálogo Maestro (Excel)")
    uploaded_master_file = st.file_uploader("Sube tu archivo Excel con productos y códigos EAN", type=["xlsx", "xls"], key="upload_master_excel")
    
    if uploaded_master_file is not None:
        try:
            df_upload = pd.read_excel(uploaded_master_file)
            st.write("Vista previa del archivo cargado:", df_upload.head(3))
            
            col_desc_opt = st.selectbox("Selecciona la columna de Nombre/Descripción:", df_upload.columns)
            col_ean_opt = st.selectbox("Selecciona la columna de Código EAN/Barras:", df_upload.columns)
            
            if st.button("📥 Importar y Fusionar al Maestro Oficial"):
                nuevos_agregados = 0
                for _, row in df_upload.iterrows():
                    p_desc = str(row[col_desc_opt]).upper().strip()
                    p_ean = clean_ean_code(row[col_ean_opt])
                    if p_desc and p_desc != "NAN" and p_ean != "S/C":
                        master_dict[p_desc] = p_ean
                        nuevos_agregados += 1
                
                save_json_file(MASTER_CATALOG_FILE, master_dict)
                st.success(f"¡Se han importado y guardado {nuevos_agregados} productos exitosamente en el maestro oficial!")
                time.sleep(0.8)
                st.rerun()
        except Exception as e:
            st.error(f"Error al leer el archivo Excel: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # AGREGAR INDIVIDUAL
    # ==========================================
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### ➕ Agregar Producto Individual Manualmente")
    with st.form("form_agregar_maestro_individual"):
        col_fm1, col_fm2 = st.columns([2, 1])
        with col_fm1:
            nuevo_nombre_prod = st.text_input("Nombre / Descripción Oficial del Producto", placeholder="EJ. ANTIOQUEÑO SIN AZUCAR TV 24% 750ML...")
        with col_fm2:
            nuevo_codigo_prod = st.text_input("Código EAN / SAP Oficial", placeholder="EJ. Código de barras real...")
        
        btn_submit_maestro = st.form_submit_button("💾 Guardar Producto en Archivo Maestro Oficial")
        if btn_submit_maestro:
            clean_n = nuevo_nombre_prod.upper().strip()
            clean_c = clean_ean_code(nuevo_codigo_prod)
            if clean_n and clean_c != "S/C":
                master_dict[clean_n] = clean_c
                save_json_file(MASTER_CATALOG_FILE, master_dict)
                st.success(f"✅ ¡Producto **{clean_n}** guardado exitosamente en el maestro oficial con el código **{clean_c}**!")
                time.sleep(0.5)
                st.rerun()
            else:
                st.error("⚠️ Por favor ingresa un nombre de producto válido y un código EAN/SAP correcto (de 7 a 14 dígitos).")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    if master_dict:
        st.markdown(f"### 📋 Productos Registrados en el Archivo Maestro ({len(master_dict):,} registros)")
        df_show = pd.DataFrame([{"Producto / Descripción": k, "Código EAN/SAP Oficial": v} for k, v in master_dict.items()])
        st.dataframe(df_show, use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ El archivo maestro está actualmente vacío. Agrega productos manualmente o cárgalos mediante Excel.")

# ==========================================
# MÓDULO 3: GESTIONAR PROVEEDORES
# ==========================================
elif menu_opcion == "🏢 Gestionar Proveedores":
    st.markdown("<h2>🏢 Configuración de Perfiles Independientes por Proveedor</h2>", unsafe_allow_html=True)
    st.markdown("---")
    supps = st.session_state["supplier_memory"]
    if not isinstance(supps, dict): supps = {}
    for p_name, p_data in list(supps.items()):
        with st.expander(f"🏢 Perfil Proveedor: {p_name}"):
            st.write(p_data)
