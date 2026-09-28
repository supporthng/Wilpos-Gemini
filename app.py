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
    page_title="WilPOS - Facturas Multi-Página", 
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

# Configuración de Clave API (Prioriza clave de pago)
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

if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    if not loaded_supps:
        loaded_supps = {
            "ALVAREZ & SANCHEZ": {"nombre": "ALVAREZ & SANCHEZ", "notas_formato": "Desglose estándar en columna tamaño (ej: 12/75 CL)."},
            "GONZALEZ CUESTA": {"nombre": "GONZALEZ CUESTA", "notas_formato": "Corporativo."}
        }
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE, "dict")

def safe_float(val, default=0.0):
    try: return float(val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): 
    return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    
    s_val = re.sub(r'\D', '', s_val)
    
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"] or len(s_val) < 8 or len(s_val) > 14:
        return "S/C"
        
    return str(s_val)

def limpiar_nombre_producto(descripcion_raw, tamano_raw):
    desc = str(descripcion_raw).upper().strip()
    desc = re.sub(r'\b(CAJA|CAJ|BOT|UNIDAD|PZA)\b', '', desc)
    desc = re.sub(r'\s+', ' ', desc).strip()
    
    tam = str(tamano_raw).upper().strip()
    m_medida = re.search(r'(\d+\s*(?:CL|ML|L|LT|G|KG))', tam)
    medida_limpia = m_medida.group(1) if m_medida else ""
    
    if medida_limpia and medida_limpia not in desc:
        return f"{desc} {medida_limpia}".strip()
    
    return desc

def parse_empaque(tamano_txt="", unidad_txt="", descripcion_txt=""):
    combined_tamano = str(tamano_txt).upper()
    m_slash = re.search(r'\b(48|24|16|12|6|10|20|30|1)\s*/', combined_tamano)
    if m_slash: 
        return int(m_slash.group(1))

    combined = f"{str(unidad_txt)} {str(descripcion_txt)}".upper()
    m_caja = re.search(r'(?:CAJA|CAJ|PAQ|PACK|BLISTER)[^\d]*(\d+)', combined)
    if m_caja:
        val = int(m_caja.group(1))
        if 1 < val <= 120: return val

    if any(w in unidad_txt.upper() for w in ["BOT", "UNIDAD", "PZA"]) and not re.search(r'\d+', unidad_txt):
        return 1
    return 1

def buscar_en_catalogo_maestro(nombre_producto):
    master_dict = st.session_state.get("master_catalog", {})
    if not master_dict: return "S/C"
    
    n_upper = str(nombre_producto).upper().strip()
    if n_upper in master_dict:
        return clean_ean_code(master_dict[n_upper])
        
    query_words = [w for w in re.findall(r'\w+', n_upper) if len(w) > 2]
    if not query_words: return "S/C"
    
    best_code = "S/C"
    max_matches = 0
    for m_name, m_code in master_dict.items():
        m_upper = str(m_name).upper()
        matches = sum(1 for qw in query_words if qw in m_upper)
        if matches >= 2 and matches > max_matches:
            max_matches = matches
            best_code = clean_ean_code(m_code)
            
    return best_code

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS System</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

# ==========================================
# MÓDULO 1: PROCESAR FACTURA
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador Inteligente de Facturas (Multi-Página)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube facturas de una o varias páginas. El sistema detecta la paginación y cruza con tu Catálogo Maestro.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""
    if "paginacion_detectada" not in st.session_state: st.session_state["paginacion_detectada"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    lista_proveedores = ["🔍 Detección Automática (Nuevo Proveedor)"] + list(st.session_state["supplier_memory"].keys())
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        prov_seleccionado = st.selectbox("🏢 Selecciona el Proveedor", lista_proveedores)
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu factura (PDF multi-página o Imagen)", type=["pdf", "png", "jpg", "jpeg"])
    
    if archivo_subido is not None:
        if st.button("🚀 Procesar Factura y Validar EAN"):
            with st.spinner("🚀 Analizando documento, paginación y extrayendo ítems..."):
                try:
                    if not gemini_key:
                        raise ValueError("No se encontró ninguna clave de API configurada.")

                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in f_type.lower() else Image.open(io.BytesIO(file_bytes))

                    prov_instruccion = f"El proveedor seleccionado es '{prov_seleccionado}'." if prov_seleccionado != "🔍 Detección Automática (Nuevo Proveedor)" else "Identifica el nombre comercial del proveedor emisor en este documento."
                    
                    prompt_unificado = (
                        f"{prov_instruccion} "
                        "Analiza el documento completo de principio a fin. "
                        "1. Detecta la información de paginación impresa en el documento (ejemplo: '1 de 2', 'Pág: 1 / 2', o si es una página única indícalo como '1 de 1'). "
                        "2. Extrae ABSOLUTAMENTE TODOS LOS RENGLONES Y PRODUCTOS de esta vista/página, sin omitir ninguno. "
                        "Para cada renglón extrae: "
                        "- 'descripcion': nombre limpio del producto. "
                        "- 'tamano': contenido de la columna tamaño (ej: '12/75 CL.'). "
                        "- 'codigo_factura': código de barras o código interno. "
                        "- 'cantidad': cantidad comprada. "
                        "- 'unidad': unidad de empaque (ej: 'CAJA'). "
                        "- 'valor_con_itbis': monto total incluyendo ITBIS. "
                        "Devuelve un JSON puro con esta estructura exacta: "
                        '{"paginacion": "1 de 2", "proveedor_detectado": "NOMBRE_PROVEEDOR", "items": [{"descripcion": "...", "tamano": "...", "codigo_factura": "...", "cantidad": 1.0, "unidad": "...", "valor_con_itbis": 0.0}]}. '
                        "Respuesta JSON pura."
                    )

                    archivo_subido.seek(0)
                    response = model.generate_content([image_input, prompt_unificado])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("```"): raw_text = raw_text[:-3]
                    
                    parsed_json = json.loads(raw_text.strip())
                    
                    prov_a_usar = prov_seleccionado
                    if prov_seleccionado == "🔍 Detección Automática (Nuevo Proveedor)":
                        prov_a_usar = str(parsed_json.get("proveedor_detectado") or "PROVEEDOR NUEVO").upper().strip()
                        supps = st.session_state["supplier_memory"]
                        if prov_a_usar not in supps:
                            supps[prov_a_usar] = {"nombre": prov_a_usar, "notas_formato": "Auto-registrado."}
                            st.session_state["supplier_memory"] = supps
                            save_json_file(SUPPLIER_MEMORY_FILE, supps)

                    st.session_state["factura_data"] = parsed_json
                    st.session_state["prov_activo"] = prov_a_usar
                    st.session_state["paginacion_detectada"] = parsed_json.get("paginacion", "Página Única (1 de 1)")
                    
                    st.success(f"✅ ¡Documento procesado con éxito! Paginación detectada: **{st.session_state['paginacion_detectada']}**")
                except Exception as e:
                    st.error(f"⚠️ Error al procesar: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state["factura_data"] is not None:
        items = st.session_state["factura_data"].get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")
        pag_info = st.session_state.get("paginacion_detectada", "1 de 1")
        
        if items:
            st.markdown(f"### 📊 Paginación: **Página {pag_info}** | Proveedor: {prov_actual} | Ítems: {len(items)}")
            preview_rows = []
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            for idx, item in enumerate(items, start=1):
                raw_desc = item.get("descripcion", "")
                raw_tam = item.get("tamano", "")
                
                nombre_completo = limpiar_nombre_producto(raw_desc, raw_tam)
                
                m_med = re.search(r'(\d+\s*(?:CL|ML|L|LT|G|KG))', str(raw_tam).upper())
                presentacion_limpia = m_med.group(1) if m_med else str(raw_tam).strip()

                cod_factura_limpio = clean_ean_code(item.get("codigo_factura"))
                if cod_factura_limpio != "S/C":
                    codigo_final = cod_factura_limpio 
                else:
                    codigo_final = buscar_en_catalogo_maestro(nombre_completo) 

                cant_compra = safe_float(item.get("cantidad"), 1.0)
                unidad = str(item.get("unidad", ""))
                val_con_itbis = safe_float(item.get("valor_con_itbis"), 0.0)

                empaque = parse_empaque(raw_tam, unidad, raw_desc)
                total_unidades = int(cant_compra * empaque)
                
                costo_sin_itbis_total = val_con_itbis / 1.18 if val_con_itbis > 0 else 0.0
                costo_unitario_real = round(costo_sin_itbis_total / total_unidades, 2) if total_unidades > 0 else 0.0

                if costo_unitario_real > 0:
                    precio_con_utilidad = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(precio_con_utilidad * 1.18)
                else:
                    precio_venta = 0.0

                preview_rows.append({
                    "No.": idx, 
                    "Producto": nombre_completo, 
                    "Código EAN Asignado": codigo_final,
                    "Cant. Compra": cant_compra,
                    "Empaque": empaque,
                    "Total Unidades": total_unidades,
                    "Costo Unit. Sin ITBIS": costo_unitario_real, 
                    "Precio Venta": precio_venta
                })

                ws.append([
                    nombre_completo, presentacion_limpia, str(codigo_final), prov_actual, "producto",
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
    st.markdown("<h2>📁 Gestión del Catálogo Maestro de Productos</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu archivo Excel masivo o registra productos de forma manual uno a uno.</p>", unsafe_allow_html=True)
    st.markdown("---")

    tab_masivo, tab_manual = st.tabs(["📂 Carga Masiva (Excel)", "➕ Agregar Producto Manualmente"])

    with tab_masivo:
        master_file = st.file_uploader("📂 Sube tu Catálogo Maestro (Excel)", type=["xlsx"])
        if master_file is not None:
            df_master = pd.read_excel(master_file, dtype=str)
            cols = df_master.columns.tolist()
            col_c1, col_c2 = st.columns(2)
            with col_c1: col_name = st.selectbox("Columna con Nombre del Producto", cols)
            with col_c2: col_code = st.selectbox("Columna con Código de Barra EAN", cols)
            if st.button("🔄 Guardar Catálogo Masivo"):
                temp_dict = st.session_state["master_catalog"]
                count = 0
                for _, row in df_master.iterrows():
                    p_name = str(row[col_name]).strip().upper()
                    p_code = clean_ean_code(row[col_code])
                    if p_name and p_code != "S/C":
                        temp_dict[p_name] = p_code
                        count += 1
                st.session_state["master_catalog"] = temp_dict
                save_json_file(MASTER_CATALOG_FILE, temp_dict)
                st.success(f"¡Catálogo actualizado con éxito! Se cargaron **{count}** productos.")

    with tab_manual:
        st.markdown("### ✍️ Registrar Producto Individual")
        with st.form("form_nuevo_producto"):
            col_m1, col_m2 = st.columns([2, 1])
            with col_m1:
                input_nombre = st.text_input("Nombre y Presentación del Producto (Ej: VINO SANTA HELENA 750 ML)")
            with col_m2:
                input_codigo = st.text_input("Código de Barra EAN (8 a 14 dígitos)")
            
            btn_guardar_manual = st.form_submit_button("💾 Guardar en Catálogo Maestro")
            
            if btn_guardar_manual:
                n_limpio = str(input_nombre).strip().upper()
                c_limpio = clean_ean_code(input_codigo)
                
                if not n_limpio:
                    st.error("⚠️ Debes ingresar el nombre del producto.")
                elif c_limpio == "S/C":
                    st.error("⚠️ El código EAN ingresado no es válido (debe ser un código de barras estándar de 8 a 14 dígitos).")
                else:
                    master_dict = st.session_state["master_catalog"]
                    master_dict[n_limpio] = c_limpio
                    st.session_state["master_catalog"] = master_dict
                    save_json_file(MASTER_CATALOG_FILE, master_dict)
                    st.success(f"✅ ¡Producto guardado con éxito! **{n_limpio}** -> EAN: **{c_limpio}**")

    master_data = st.session_state.get("master_catalog", {})
    if master_data:
        st.markdown(f"### 📋 Productos en Catálogo Maestro ({len(master_data):,} registros)")
        df_show = pd.DataFrame([{"Producto": k, "Código EAN Oficial": v} for k, v in master_data.items()])
        st.dataframe(df_show, use_container_width=True, hide_index=True)

# ==========================================
# MÓDULO 3: GESTIONAR PROVEEDORES
# ==========================================
elif menu_opcion == "🏢 Gestionar Proveedores":
    st.markdown("<h2>🏢 Perfiles de Proveedores Memorizados</h2>", unsafe_allow_html=True)
    st.markdown("---")
    supps = st.session_state["supplier_memory"]
    for p_name, p_data in supps.items():
        with st.expander(f"🏢 {p_name}"):
            st.write(f"**Nombre:** {p_data.get('nombre', p_name)}")
            st.write(f"**Detalles:** {p_data.get('notas_formato', 'N/A')}")
