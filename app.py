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
# GESTIÓN DE PERFILES SEPARADOS Y CATÁLOGO
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "PRICESMART": {
            "nombre": "PRICESMART",
            "tipo_formato": "factura_tique_unidades",
            "instruccion_prompt": "Analiza este comprobante de PRICESMART renglón por renglón. Extrae 'descripcion', 'tamano' (ej. 12OZ), 'cantidad', 'unidad' (ej. EA, 1UN), 'precio_unitario' y 'monto_neto'."
        },
        "CENTRO DE DISTRIBUCION CRISTIAN": {
            "nombre": "CENTRO DE DISTRIBUCION CRISTIAN",
            "tipo_formato": "pos_cajas_unidades",
            "instruccion_prompt": "Analiza este documento de CENTRO DE DISTRIBUCION CRISTIAN (CDC) renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_desglose_descuentos",
            "instruccion_prompt": "Analiza esta factura de ALVAREZ & SANCHEZ renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "ELÍAS DISTRIBUCIÓN": {
            "nombre": "ELÍAS DISTRIBUCIÓN",
            "tipo_formato": "factura_cajas_unidades",
            "instruccion_prompt": "Analiza esta factura de ELÍAS DISTRIBUCIÓN renglón por renglón. Extrae 'descripcion', 'unidad', 'cantidad', 'precio_unitario' y 'monto_neto'."
        },
        "EL CATADOR": {
            "nombre": "EL CATADOR",
            "tipo_formato": "factura_cajas_descuento",
            "instruccion_prompt": "Analiza esta factura de EL CATADOR renglón por renglón. Extrae 'descripcion', 'unidad', 'cantidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
        },
        "CND / BEES": {
            "nombre": "CND / BEES",
            "tipo_formato": "tique_doble_linea_blindado",
            "instruccion_prompt": "Analiza este tique de CND / BEES renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
        }
    }
    if not loaded_supps or "PRICESMART" not in loaded_supps:
        loaded_supps.update(default_profiles)
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    loaded_master = load_json_file(MASTER_CATALOG_FILE, "dict")
    if not loaded_master:
        base_defaults = {
            "ANTIOQUEÑO TAPA ROJA 750 ML": "7702131234567",
            "ANTIOQUEÑO TAPA AZUL 750 ML": "7702131234574",
            "ANTIOQUEÑO TAPA VERDE 750 ML": "7702131234581",
            "OLD PARR 12 AÑOS 750ML": "7804300120986",
            "JOHNNIE WALKER BLUE LABEL 750ML": "5000267022108",
            "TEQUILA DON JULIO REPOSADO 750ML": "7501035602409",
            "CANADA DRY GINGER ALE 12OZ": "078114031201"
        }
        loaded_master = base_defaults
        save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

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
    
    if "PRICESMART" in prov_up:
        t_norm = re.sub(r'^\d+\s+', '', t_norm)
        t_norm = re.sub(r'^MS\s*', '', t_norm)
    
    # Conversiones de marcas
    t_norm = re.sub(r'\bJW\b', 'JOHNNIE WALKER', t_norm)
    t_norm = re.sub(r'\bDJ\b', 'DON JULIO', t_norm)
    if "OLD PARR" in t_norm:
        t_norm = "OLD PARR 12 AÑOS"
    
    combined_raw = f"{t_norm} {t_tam}"
    
    m_med = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|CL))', combined_raw)
    presentacion = m_med.group(1).replace(" ", "") if m_med else ("750ML" if "OLD PARR" in t_norm else "UN")
    if "75CL" in presentacion:
        presentacion = "750ML"

    if "ANTIOQ" in combined_raw:
        desc_limpia = f"ANTIOQUEÑO TAPA ROJA {presentacion}"
    elif "OLD PARR" in t_norm:
        desc_limpia = f"OLD PARR 12 AÑOS {presentacion}"
    else:
        desc_limpia = re.sub(r'\b(CAJA\s*\d*|CJ\s*\d*\s*BOT|\d+\s*X\s*\d+\s*(?:ML|CL|L)|\d+\s+1UN|1UN|EA|750\s*ML|75\s*CL|750ML|75CL|\d+OZ)\b', '', t_norm)
        desc_limpia = re.sub(r'\s+', ' ', desc_limpia).strip()
        if presentacion != "UN" and presentacion not in desc_limpia:
            desc_limpia = f"{desc_limpia} {presentacion}".strip()
    
    return desc_limpia, presentacion

def parse_empaque_proveedor(proveedor_nombre, unidad_txt="", tamano_txt="", descripcion_txt=""):
    combined = normalizar_texto(f"{unidad_txt} {tamano_txt} {descripcion_txt}")
    prov_up = normalizar_texto(proveedor_nombre)
    
    if "PRICESMART" in prov_up:
        m_pack = re.search(r'\b(12|24|6|48)\b', combined)
        if m_pack:
            val = int(m_pack.group(1))
            if val > 1: return val
        if "EA" in combined or "1UN" in combined or ("UN" in combined and "12" not in combined):
            return 1

    m_caj_num = re.search(r'(?:CAJA|CJ|BOX)[\s\-]*(\d+)', combined)
    if m_caj_num:
        val = int(m_caj_num.group(1))
        if val > 0: return val

    if "12X" in combined or "12/" in combined or "CJ12" in combined:
        return 12
    if "24X" in combined or "24/" in combined or "CJ24" in combined:
        return 24

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_norm = normalizar_texto(nombre_producto)
    p_norm = normalizar_texto(presentacion)
    combined_query = normalizar_texto(f"{n_norm} {p_norm}")

    master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")
    master_norm = {normalizar_texto(k): v for k, v in master_dict.items()}

    if combined_query in master_norm:
        return clean_ean_code(master_norm[combined_query])
    if n_norm in master_norm:
        return clean_ean_code(master_norm[n_norm])

    tokens_query = set(re.findall(r'\b[A-Z0-9]+\b', combined_query))
    tokens_query = {t for t in tokens_query if len(t) > 1 and t not in {"ML", "CL", "L", "OZ", "CON", "SIN", "EA", "UN"}}

    if not tokens_query:
        return "S/C"

    mejor_codigo = "S/C"
    max_coincidentes = 0

    for m_key, m_code in master_norm.items():
        tokens_master = set(re.findall(r'\b[A-Z0-9]+\b', m_key))
        comunes = tokens_query.intersection(tokens_master)
        score = len(comunes)
        
        if score > max_coincidentes and score >= 2:
            max_coincidentes = score
            mejor_codigo = clean_ean_code(m_code)

    return mejor_codigo if mejor_codigo != "S/C" else "S/C"

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS Multi-Proveedor</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Extracción de Presentaciones en OZ Activa</p>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: PROCESAR FACTURA
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador con Reconocimiento de Formatos en Onzas (OZ)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu factura o tique. El sistema extrae correctamente tamaños como 12OZ y empareja con el catálogo.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""
    if "paginacion_detectada" not in st.session_state: st.session_state["paginacion_detectada"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        st.info("💡 Sube tu factura o tique (PDF o Imagen).")
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu factura o tique (PDF multi-página o Imagen)", type=["pdf", "png", "jpg", "jpeg"])
    
    if archivo_subido is not None:
        if st.button("🚀 Detectar Proveedor y Procesar con Precisión"):
            with st.spinner("🔍 Analizando documento con soporte para onzas (OZ)..."):
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
                            "instruccion_prompt": f"Analiza esta factura de {prov_encontrado} renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'."
                        }
                        st.session_state["supplier_memory"] = supp_mem
                        save_json_file(SUPPLIER_MEMORY_FILE, supp_mem)

                    prov_dict_data = supp_mem.get(prov_encontrado, {})
                    if not isinstance(prov_dict_data, dict): prov_dict_data = {}
                    instruccion_proveedor = prov_dict_data.get("instruccion_prompt", "Extrae todos los ítems.")

                    prompt_unificado = (
                        f"Estás procesando un tique o factura del proveedor: '{prov_encontrado}'. "
                        f"Instrucción específica de su perfil: {instruccion_proveedor} "
                        "Extrae 'descripcion', 'tamano' (ej. 12OZ), 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'. "
                        "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                        '{"paginacion": "1 de 1", "proveedor_detectado": "' + prov_encontrado + '", "subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"descripcion": "...", "tamano": "12OZ", "cantidad": 1.0, "unidad": "EA", "precio_unitario": 4.58, "descuento_porcentaje": 0.0, "monto_neto": 4.58}]}. '
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

        total_importe_neto = sum(safe_float(i.get("monto_neto") or i.get("precio_unitario")) for i in items)
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
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo Unitario', 'Stock (Unidades)', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

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

                p_unit_extraido = safe_float(item.get("precio_unitario"), 0.0)
                desc_pct = safe_float(item.get("descuento_porcentaje"), 0.0)
                monto_neto_linea = safe_float(item.get("monto_neto"), 0.0)

                if p_unit_extraido > 0:
                    precio_neto_item = p_unit_extraido * (1 - (desc_pct / 100.0))
                    costo_unitario_real = round(precio_neto_item / empaque, 2) if empaque > 1 else precio_neto_item
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
                    st.markdown(f"⚠️ **{nombre_limpio} ({presentacion_limpia})** sin código EAN asignado (requiere validación).")
                    
                    col_c1, col_c2, col_c3 = st.columns([2, 1, 1])
                    with col_c1:
                        sel_maestro = st.selectbox(
                            f"Seleccionar del Catálogo Maestro",
                            ["-- Buscar en Maestro --"] + nombres_maestro_lista,
                            key=f"sel_maestro_{idx}_{nombre_limpio}_{idx}"
                        )
                        if sel_maestro != "-- Buscar en Maestro --":
                            display_codigo = master_dict[sel_maestro]
                    with col_c2:
                        codigo_manual_input = st.text_input("Ingresar código manual", key=f"manual_input_{idx}_{nombre_limpio}_{idx}", placeholder="Ej. 078114031201...")
                        if codigo_manual_input and len(codigo_manual_input.strip()) >= 7:
                            clean_m = clean_ean_code(codigo_manual_input)
                            if clean_m != "S/C":
                                display_codigo = clean_m
                    with col_c3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        query_busqueda = f"EAN barcode {nombre_limpio} {presentacion_limpia}".replace(" ", "+")
                        url_busqueda = f"https://www.google.com/search?q={query_busqueda}"
                        st.markdown(f"[🌐 Buscar en la Web]({url_busqueda})", unsafe_allow_html=True)

                    if display_codigo != "S/C":
                        if st.button("✅ Guardar y Usar este Código", key=f"btn_conf_{idx}_{nombre_limpio}"):
                            st.session_state["codigos_manuales_sesion"][nombre_limpio] = display_codigo
                            master_dict[nombre_limpio] = display_codigo
                            save_json_file(MASTER_CATALOG_FILE, master_dict)
                            st.success(f"¡Código {display_codigo} guardado y aplicado!")
                            time.sleep(0.5)
                            st.rerun()

                preview_rows.append({
                    "No.": idx, "Producto": nombre_limpio, "Presentación": presentacion_limpia, 
                    "Código EAN": display_codigo, "Cant. Compra": cant_compra, "Empaque": empaque, 
                    "Stock Unidades": total_unidades, "Costo Unit. Real": costo_unitario_real, "Precio Venta": precio_venta
                })

                ws.append([
                    nombre_limpio, presentacion_limpia, str(display_codigo), prov_actual, "producto",
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
    st.markdown("<h2>📁 Gestión, Carga y Limpieza del Archivo Maestro EAN</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu archivo Excel con el catálogo maestro, limpia registros duplicados o agrega productos individualmente.</p>", unsafe_allow_html=True)
    st.markdown("---")

    master_dict = load_json_file(MASTER_CATALOG_FILE, "dict")

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown("### 📤 Cargar Archivo Masivo al Catálogo Maestro")
    st.markdown("<p style='font-size: 0.85rem; color: #64748b;'>Sube tu Excel o CSV con tu listado de productos y códigos oficiales. El sistema fusionará todo automáticamente.</p>", unsafe_allow_html=True)
    
    archivo_maestro_subido = st.file_uploader("Sube tu archivo Excel/CSV del Catálogo Maestro", type=["xlsx", "xls", "csv"], key="uploader_maestro_masivo")
    if archivo_maestro_subido is not None:
        if st.button("📥 Procesar, Fusionar y Guardar en Archivo Maestro"):
            try:
                if archivo_maestro_subido.name.endswith('.csv'):
                    df_m = pd.read_csv(archivo_maestro_subido)
                else:
                    df_m = pd.read_excel(archivo_maestro_subido)
                
                cols_up = [str(c).upper().strip() for c in df_m.columns]
                col_nombre_idx = next((i for i, c in enumerate(cols_up) if any(k in c for k in ['NOMBRE', 'DESCRIPCION', 'PRODUCTO'])), 0)
                col_codigo_idx = next((i for i, c in enumerate(cols_up) if any(k in c for k in ['CODIGO', 'EAN', 'BARRA', 'SAP'])), 1)

                nuevos_cargados = 0
                for _, row in df_m.iterrows():
                    p_nombre = str(row.iloc[col_nombre_idx]).upper().strip()
                    p_codigo = str(row.iloc[col_codigo_idx]).strip()
                    clean_c = clean_ean_code(p_codigo)
                    if p_nombre and p_nombre != "NAN" and clean_c != "S/C":
                        master_dict[p_nombre] = clean_c
                        nuevos_cargados += 1
                
                save_json_file(MASTER_CATALOG_FILE, master_dict)
                st.success(f"✅ ¡Se cargaron y fusionaron **{nuevos_cargados}** productos exitosamente al archivo maestro!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Error al procesar el archivo: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### 🧹 Herramientas de Mantenimiento")
    col_l1, col_l2 = st.columns([1, 2])
    with col_l1:
        if st.button("🗑️ Resetear / Limpiar Archivo Maestro"):
            base_inicial = {
                "ANTIOQUEÑO TAPA ROJA 750 ML": "7702131234567",
                "ANTIOQUEÑO TAPA AZUL 750 ML": "7702131234574",
                "ANTIOQUEÑO TAPA VERDE 750 ML": "7702131234581",
                "OLD PARR 12 AÑOS 750ML": "7804300120986",
                "CANADA DRY GINGER ALE 12OZ": "078114031201"
            }
            save_json_file(MASTER_CATALOG_FILE, base_inicial)
            st.success("¡Archivo maestro restablecido y limpiado correctamente!")
            time.sleep(0.5)
            st.rerun()
    with col_l2:
        st.info("ℹ️ Borra entradas corruptas y restaura los productos base esenciales.")

    st.markdown("---")

    with st.expander("➕ Agregar Producto Individual Manualmente", expanded=False):
        with st.form("form_agregar_maestro"):
            col_m1, col_m2 = st.columns([2, 1])
            with col_m1:
                nuevo_prod_nombre = st.text_input("Nombre / Descripción Oficial (Ej: CANADA DRY GINGER ALE 12OZ)")
            with col_m2:
                nuevo_prod_codigo = st.text_input("Código EAN / SAP Oficial (Ej: 078114031201)")
            
            btn_guardar_maestro = st.form_submit_button("💾 Guardar en Archivo Maestro")
            if btn_guardar_maestro:
                clean_name = nuevo_prod_nombre.upper().strip()
                clean_code = clean_ean_code(nuevo_prod_codigo)
                if clean_name and clean_code != "S/C":
                    master_dict[clean_name] = clean_code
                    save_json_file(MASTER_CATALOG_FILE, master_dict)
                    st.success(f"✅ ¡Producto **{clean_name}** guardado con éxito!")
                    st.rerun()
                else:
                    st.error("⚠️ Por favor ingresa un nombre válido y un código EAN/SAP correcto.")

    if master_dict:
        st.markdown(f"### 📋 Productos Registrados en el Archivo Maestro ({len(master_dict):,} registros limpios)")
        df_show = pd.DataFrame([{"Producto / Descripción": k, "Código EAN/SAP Oficial": v} for k, v in master_dict.items()])
        st.dataframe(df_show, use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 3: GESTIONAR PROVEEDORES
# ==========================================
elif menu_opcion == "🏢 Gestionar Proveedores":
    st.markdown("<h2>🏢 Configuración de Perfiles Independientes por Proveedor</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Administra, edita o crea perfiles separados para cada proveedor.</p>", unsafe_allow_html=True)
    st.markdown("---")

    supps = st.session_state["supplier_memory"]
    if not isinstance(supps, dict): supps = {}
    
    for p_name, p_data in list(supps.items()):
        if not isinstance(p_data, dict): p_data = {}
        with st.expander(f"🏢 Perfil Proveedor: {p_name}"):
            with st.form(f"form_prov_{p_name}"):
                nuevo_nombre = st.text_input("Nombre del Proveedor", value=p_data.get("nombre", p_name))
                nueva_instruccion = st.text_area("Instrucción / Prompt de Formato Exclusivo", value=p_data.get("instruccion_prompt", ""), height=120)
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    btn_guardar = st.form_submit_button("💾 Guardar Cambios del Perfil")
                with col_btn2:
                    btn_eliminar = st.form_submit_button("🗑️ Eliminar Perfil")
                
                if btn_guardar:
                    supps[p_name] = {"nombre": nuevo_nombre, "instruccion_prompt": nueva_instruccion}
                    if nuevo_nombre != p_name:
                        supps[nuevo_nombre] = supps.pop(p_name)
                    st.session_state["supplier_memory"] = supps
                    save_json_file(SUPPLIER_MEMORY_FILE, supps)
                    st.success(f"✅ ¡Perfil de **{nuevo_nombre}** actualizado con éxito!")
                    st.rerun()
                    
                if btn_eliminar:
                    if p_name in supps:
                        supps.pop(p_name)
                        st.session_state["supplier_memory"] = supps
                        save_json_file(SUPPLIER_MEMORY_FILE, supps)
                        st.warning(f"⚠️ Perfil de {p_name} eliminado.")
                        st.rerun()

    with st.expander("➕ Crear Nuevo Perfil de Proveedor Independiente"):
        with st.form("form_nuevo_proveedor_manual"):
            n_prov = st.text_input("Nombre del Proveedor (Ej: CASA BRUGAL)")
            n_inst = st.text_area("Instrucción de Formato para este Proveedor", value="Analiza la factura de este proveedor renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario', 'descuento_porcentaje' y 'monto_neto'.")
            btn_crear = st.form_submit_button("Crear Nuevo Perfil")
            if btn_crear:
                clean_p = n_prov.upper().strip()
                if clean_p:
                    supps[clean_p] = {"nombre": clean_p, "tipo_formato": "personalizado", "instruccion_prompt": n_inst}
                    st.session_state["supplier_memory"] = supps
                    save_json_file(SUPPLIER_MEMORY_FILE, supps)
                    st.success(f"✅ ¡Perfil independiente creado para {clean_p}!")
                    st.rerun()
