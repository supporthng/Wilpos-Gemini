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

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else os.getcwd()
SUPPLIER_MEMORY_FILE = os.path.join(BASE_DIR, "proveedores_formatos_memoria.json")
MASTER_CATALOG_FILE = os.path.join(BASE_DIR, "catalogo_maestro_sistema.json")

def load_json_file(filepath, default_type="dict"):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if default_type == "dict":
                    if isinstance(data, dict):
                        return {str(k).upper().strip(): str(v) for k, v in data.items() if k}
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
        return True
    except Exception:
        return False

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
    combined_raw = f"{t_norm} {t_tam}"
    
    presentacion = "330ML"
    if "12 OZ" in combined_raw or "12ONZ" in combined_raw: presentacion = "330ML"
    elif "290" in combined_raw: presentacion = "290ML"
    elif "500" in combined_raw: presentacion = "500ML"
    elif "750" in combined_raw: presentacion = "750ML"
    elif "1000" in combined_raw or "1 LT" in combined_raw: presentacion = "1000ML"

    desc_limpia = re.sub(r'\b\d{4,5}\b', '', t_norm)
    desc_limpia = re.sub(r'\b(CAJA|CJ|\d+/\d+|\d+\s*X\s*\d+|12\s*OZ|290\s*ML|330\s*ML|500\s*ML|750\s*ML|1000\s*ML|LATA|LATAS|BOT|BOTELLA)\b', '', desc_limpia)
    desc_limpia = re.sub(r'\s+', ' ', desc_limpia).strip()
    
    if presentacion not in desc_limpia:
        desc_limpia = f"{desc_limpia} {presentacion}".strip()
        
    return desc_limpia, presentacion

def parse_empaque_proveedor(proveedor_nombre, unidad_txt="", tamano_txt="", descripcion_txt=""):
    unidad_norm = normalizar_texto(unidad_txt)
    combined = normalizar_texto(f"{unidad_txt} {tamano_txt} {descripcion_txt}")
    
    m_cj = re.search(r'CJ\s*(\d+)', combined)
    if m_cj:
        val = int(m_cj.group(1))
        if val > 1: return val

    if "LATA" in combined or "ULTRA" in combined or "CJ" in combined:
        return 24

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_norm = normalizar_texto(nombre_producto)
    master_dict = st.session_state.get("master_catalog", load_json_file(MASTER_CATALOG_FILE, "dict"))
    if not master_dict:
        return "S/C"

    master_norm = {normalizar_texto(k): clean_ean_code(v) for k, v in master_dict.items() if k}

    query_completa = f"{n_norm}".strip()
    for m_key_up, m_code in master_norm.items():
        if m_key_up == query_completa:
            return m_code

    tokens_query = set(re.findall(r'\b[A-Z0-9%]+\b', n_norm))
    tokens_query = {t for t in tokens_query if len(t) > 1 and t not in {"ML", "LT", "L", "OZ", "BOT", "UN", "PC"}}
    if not tokens_query:
        return "S/C"

    mejor_codigo = "S/C"
    max_coincidentes = 0

    for m_key_up, m_code in master_norm.items():
        tokens_master = set(re.findall(r'\b[A-Z0-9%]+\b', m_key_up))
        comunes = tokens_query.intersection(tokens_master)
        score = len(comunes)

        if score > max_coincidentes and score >= max(1, len(tokens_query) // 2):
            max_coincidentes = score
            mejor_codigo = m_code

    return mejor_codigo

# ==========================================
# GESTIÓN DE PERFILES Y CATÁLOGO PERMANENTE
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "CERVECERIA NACIONAL DOMINICANA": {
            "nombre": "CERVECERIA NACIONAL DOMINICANA",
            "tipo_formato": "factura_tique_cnd",
            "instruccion_prompt": "Analiza esta factura renglón por renglón. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad', 'precio_unitario' y 'monto_neto' (Imp. Neto sin ITBIS)."
        }
    }
    if not loaded_supps or "CERVECERIA NACIONAL DOMINICANA" not in loaded_supps:
        loaded_supps.update(default_profiles)
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE, "dict")

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

if "paginas_procesadas_historial" not in st.session_state:
    st.session_state["paginas_procesadas_historial"] = set()

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS Multi-Proveedor</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura (Por Página)", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown(f"<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Maestro Permanente ({len(st.session_state.get('master_catalog', {})):,} ítems)</p>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: PROCESAR FACTURA POR PÁGINA
# ==========================================
if menu_opcion == "📄 Procesar Factura (Por Página)":
    st.markdown("<h2>📄 Procesador de Facturas (Multi-Proveedor)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Edita directamente 'Cant. Compra' o 'Empaque' y el stock y costos se recalcularán al instante.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1: st.info("💡 Sube tu tique o factura (PDF o imagen).")
    with col_s2: margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu archivo", type=["pdf", "png", "jpg", "jpeg"])
    
    pagina_a_procesar = 1
    if archivo_subido is not None and archivo_subido.name.endswith('.pdf'):
        try:
            import fitz
            archivo_subido.seek(0)
            doc_pdf = fitz.open(stream=archivo_subido.read(), filetype="pdf")
            total_paginas = len(doc_pdf)
            pagina_a_procesar = st.selectbox("Selecciona la página exacta:", list(range(1, total_paginas + 1)))
        except Exception:
            pagina_a_procesar = 1

    if archivo_subido is not None:
        if st.button(f"🚀 Procesar Página #{pagina_a_procesar}"):
            with st.spinner(f"🔍 Analizando página #{pagina_a_procesar}..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave API.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    
                    if "pdf" in f_type.lower():
                        import fitz
                        doc_pdf = fitz.open(stream=file_bytes, filetype="pdf")
                        page_obj = doc_pdf[pagina_a_procesar - 1]
                        pix = page_obj.get_pixmap(dpi=150)
                        target_image = Image.open(io.BytesIO(pix.tobytes("jpeg")))
                    else:
                        target_image = Image.open(io.BytesIO(file_bytes))

                    prompt_det = "Identifica el proveedor y factura. Devuelve JSON: {\"proveedor_detectado\": \"...\", \"numero_factura\": \"...\"}"
                    resp_det = model.generate_content([target_image, prompt_det])
                    raw_det = resp_det.text.strip().replace("```json", "").replace("```", "").strip()
                    det_json = json.loads(raw_det)
                    prov_nombre = str(det_json.get("proveedor_detectado", "PROVEEDOR")).upper().strip()

                    prompt_items = (
                        f"Analiza esta factura de '{prov_nombre}' renglón por renglón. "
                        "Extrae estrictamente un JSON con esta estructura exacta: "
                        '{"proveedor_detectado": "' + prov_nombre + '", "items": [{"descripcion": "...", "tamano": "...", "cantidad": 2.0, "unidad": "PC", "precio_unitario": 24512.85, "monto_neto": 24512.85}]}. '
                        "CRÍTICO: Extrae 'monto_neto' sin ITBIS. Devuelve solo JSON puro."
                    )
                    resp_items = model.generate_content([target_image, prompt_items])
                    raw_items = resp_items.text.strip().replace("```json", "").replace("```", "").strip()
                    parsed_json = json.loads(raw_items)

                    st.session_state["factura_data"] = parsed_json
                    st.session_state["prov_activo"] = prov_nombre
                    st.success(f"🎯 ¡Factura procesada con éxito! Proveedor: {prov_nombre}")
                except Exception as e:
                    st.error(f"⚠️ Error: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state["factura_data"] is not None:
        data_resp = st.session_state["factura_data"]
        items = data_resp.get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")

        if items:
            st.markdown(f"### 📋 Detalle de Renglones ({len(items)} ítems)")
            lista_renglones = []
            master_dict = st.session_state.get("master_catalog", {})
            nombres_maestro = list(master_dict.keys())

            for idx, item in enumerate(items, start=1):
                raw_desc = str(item.get("descripcion", ""))
                raw_tam = str(item.get("tamano", ""))
                unidad = str(item.get("unidad", ""))
                
                nombre_limpio, presentacion_limpia = limpiar_nombre_y_extraer_presentacion(prov_actual, raw_desc, raw_tam)
                cant_compra = safe_float(item.get("cantidad"), 1.0)
                empaque = parse_empaque_proveedor(prov_actual, unidad, raw_tam, raw_desc)

                codigo_final = buscar_en_catalogo_maestro(nombre_limpio, presentacion_limpia)
                if nombre_limpio in st.session_state["codigos_manuales_sesion"]:
                    codigo_final = st.session_state["codigos_manuales_sesion"][nombre_limpio]

                monto_neto = safe_float(
                    item.get("monto_neto") or 
                    item.get("importe_neto") or 
                    item.get("neto") or 
                    item.get("imp_neto")
                )
                precio_ext = safe_float(item.get("precio_unitario") or item.get("p_unit"))

                if monto_neto <= 0 and precio_ext > 0:
                    monto_neto = precio_ext * cant_compra

                lista_renglones.append({
                    "idx": idx, 
                    "Producto": nombre_limpio, 
                    "Presentación": presentacion_limpia,
                    "Código EAN": str(codigo_final), 
                    "Cant. Compra": cant_compra, 
                    "Empaque": empaque,
                    "_monto_base": monto_neto
                })

            df_preview = pd.DataFrame(lista_renglones)
            
            # HABILITAMOS LA EDICIÓN DE 'Cant. Compra' y 'Empaque' junto con 'Código EAN'
            df_editado = st.data_editor(
                df_preview.drop(columns=["_monto_base"]),
                use_container_width=True,
                hide_index=True,
                disabled=["idx", "Producto", "Presentación"],
                column_config={
                    "Código EAN": st.column_config.TextColumn("Código EAN", max_chars=14, validate="^\\d+$"),
                    "Cant. Compra": st.column_config.NumberColumn("Cant. Compra", min_value=0.01, step=1.0),
                    "Empaque": st.column_config.NumberColumn("Empaque", min_value=1, step=1.0)
                },
                key="editor_inventario_editable"
            )

            # RE-CÁLCULO DINÁMICO EN TIEMPO REAL SEGÚN LO QUE EDITE EL USUARIO
            lista_final_procesada = []
            for i, row in df_editado.iterrows():
                p_name = row["Producto"]
                nuevo_ean = clean_ean_code(row["Código EAN"])
                if nuevo_ean != "S/C":
                    st.session_state["codigos_manuales_sesion"][p_name] = nuevo_ean

                c_compra = safe_float(row["Cant. Compra"], 1.0)
                emp = int(safe_float(row["Empaque"], 1.0))
                if emp < 1: emp = 1

                stock_unidades = int(c_compra * emp)
                
                # Monto base extraído originalmente para este renglón
                monto_base_original = df_preview.loc[i, "_monto_base"]
                
                # Costo Unitario Real limpio sobre stock actualizado
                if monto_base_original > 0 and stock_unidades > 0:
                    costo_unitario_real = round(monto_base_original / stock_unidades, 4)
                else:
                    costo_unitario_real = 0.0

                if costo_unitario_real > 0:
                    p_util = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(p_util * 1.18)
                else:
                    precio_venta = 0.0

                lista_final_procesada.append({
                    "idx": row["idx"],
                    "Producto": p_name,
                    "Presentación": row["Presentación"],
                    "Código EAN": str(nuevo_ean),
                    "Cant. Compra": c_compra,
                    "Empaque": emp,
                    "Stock Unidades": stock_unidades,
                    "Costo Unit. Real": costo_unitario_real,
                    "Precio Venta": precio_venta
                })

            df_resultado_final = pd.DataFrame(lista_final_procesada)
            st.markdown("### 📊 Vista Previa Recalculada en Tiempo Real")
            st.dataframe(df_resultado_final, use_container_width=True, hide_index=True)

            # Exportación Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo Unitario', 'Stock (Unidades)', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            for _, row in df_resultado_final.iterrows():
                ws.append([
                    row["Producto"], row["Presentación"], str(row["Código EAN"]), prov_actual, "producto",
                    row["Precio Venta"], row["Costo Unit. Real"], row["Stock Unidades"], 0.18, "unidad", row["Empaque"]
                ])

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            st.download_button(
                label=f"📥 Descargar Excel - {prov_actual}",
                data=excel_buffer.getvalue(),
                file_name=f"Inventario_{prov_actual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==========================================
# MÓDULO 2: CATÁLOGO MAESTRO EAN (PERMANENTE)
# ==========================================
elif menu_opcion == "📁 Catálogo Maestro EAN":
    st.markdown("<h2>📁 Gestión Permanente del Archivo Maestro EAN</h2>", unsafe_allow_html=True)
    master_dict = st.session_state.get("master_catalog", load_json_file(MASTER_CATALOG_FILE, "dict"))
    if master_dict:
        st.dataframe(pd.DataFrame([{"Producto": k, "Código EAN": v} for k, v in master_dict.items()]), use_container_width=True)

# ==========================================
# MÓDULO 3: GESTIONAR PROVEEDORES
# ==========================================
elif menu_opcion == "🏢 Gestionar Proveedores":
    st.markdown("<h2>🏢 Proveedores Configurados</h2>", unsafe_allow_html=True)
    for p_name, p_data in st.session_state.get("supplier_memory", {}).items():
        st.expander(p_name).write(p_data)
