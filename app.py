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
# GESTIÓN DE PERFILES AISLADOS POR PROVEEDOR
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "CND / BEES": {
            "nombre": "CND / BEES",
            "tipo_formato": "tique_doble_linea",
            "instruccion_prompt": (
                "Analiza este tique de CND / BEES renglón por renglón. "
                "Extrae estrictamente: 'descripcion', 'tamano', 'cantidad', 'unidad' (PC/UN) "
                "y el valor monetario de la columna 'impuesto_neto'."
            )
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_desglose",
            "instruccion_prompt": (
                "Analiza esta factura de ALVAREZ & SANCHEZ renglón por renglón. "
                "Extrae estrictamente: 'descripcion', 'tamano', 'cantidad', 'unidad' (CAJ/PZA) "
                "y el valor monetario exacto de la columna 'monto_neto'."
            )
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
        "GATORADE UVA": "92736", "FOUR LOKO MARACUYA": "849806004962", "FOUR LOKO PONCHE DE FRUTAS": "849806001220",
        "FOUR LOKO GREEN": "849806001855", "FOUR LOKO PURPLE": "849806002746", "FOUR LOKO GOLD": "849806001756",
        "FOUR LOKO SANDIA": "849806001206", "FOUR LOKO WHITE": "849806005754", "ALOE PURE PLUS ORIGINAL": "8809125063011",
        "MY COCO PURE PLUS": "8809125063011",
        # Álvarez & Sánchez (Pág. 3)[cite: 3]
        "SANTA HELENA MERLOT 75 CL": "7804300120986",[cite: 3]
        "SANTA HELENA RESERVADO RED BLEND 75 CL": "7804300150082",[cite: 3]
        "SANTA HELENA SAUVIGNON BLANC 75 CL": "7804300150041",[cite: 3]
        "SANTA HELENA VINO DULCE TINTO 75 CL": "7804300149307",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 1.5 LT": "8420976010063",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 375 CL": "8420976010087",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 75 CL": "842097601070",[cite: 3]
        "SCHWEPPES AGUA TONICA 4 PACK 18 CL": "2000011980849",[cite: 3]
        "SCHWEPPES TONICA 1 LT": "2117974",[cite: 3]
        "SCHWEPPES TONICA ZERO 1 LT": "2138531",[cite: 3]
        "SELA BODEGAS RODA VINO TINTO 75 CL": "8014396003073",[cite: 3]
        "SOLAN DE CABRAS AGUA MINERAL NAT 1.5 LT": "8436538810767"[cite: 3]
    }
    for k, v in base_defaults.items():
        if k not in loaded_master: loaded_master[k] = v
    save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

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
    if 8 <= len(s_val) <= 14: return str(s_val)
    return "S/C"

def limpiar_nombre_producto(descripcion_raw, tamano_raw):
    desc = str(descripcion_raw).upper().strip()
    desc = re.sub(r'\s+\d{4,6}$', '', desc)
    desc = re.sub(r'\b(PC|UN|CAJA|CAJ|BOT|PZA)\b', '', desc)
    desc = re.sub(r'\s+', ' ', desc).strip()
    
    tam = str(tamano_raw).upper().strip()
    m_medida = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z))', tam + " " + desc)
    medida_limpia = m_medida.group(1) if m_medida else ""
    
    if medida_limpia and medida_limpia not in desc:
        return f"{desc} {medida_limpia}".strip()
    
    return desc

def parse_empaque(tamano_txt="", unidad_txt="", descripcion_txt=""):
    unidad_upper = str(unidad_txt).upper()
    combined = f"{str(tamano_txt)} {str(unidad_txt)} {str(descripcion_txt)}".upper()
    
    if "12 PZA" in combined or ("12" in combined and "PZA" in unidad_upper): return 12
    if "6 PZA" in combined or ("6" in combined and "PZA" in unidad_upper): return 6
    if "24 PZA" in combined or ("24" in combined and "PZA" in unidad_upper): return 24
    
    if "CAJ" in unidad_upper or "CAJA" in unidad_upper:
        m_pack = re.search(r'\b(24|12|6|18|4)\b', combined)
        if m_pack: return int(m_pack.group(1))
            
    if "ALOE" in combined: return 24
    if "CORONA" in combined or "MICHELOB" in combined or "BRAHMA" in combined or "PTE" in combined or "THE ONE" in combined:
        if "PC" in unidad_upper and not re.search(r'\b(6|12|16)\b', combined): return 24
    if "CLAMATO" in combined: return 12
    if "FOUR LOKO" in combined: return 12
    if "GATORADE" in combined: return 24
    if "MY COCO" in combined: return 20
    if "ENRIQUILLO" in combined: return 24
        
    m_pack = re.search(r'\b(48|24|16|12|6|10|20|30|4)\s*[/xX]', combined)
    if m_pack: return int(m_pack.group(1))

    return 12 if "SANTA HELENA" in combined or "SELA" in combined else (6 if "SANTIAGO RUIZ" in combined else 1)

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_upper = str(nombre_producto).upper().strip()
    p_upper = str(presentacion).upper().strip()
    combined_query = f"{n_upper} {p_upper}".strip()

    master_dict = st.session_state.get("master_catalog", {})

    cnd_sinonimos = {
        "PTE. LIGHT HU 22OZ": "70601561", "PRESIDENTE LIGHT HU 22OZ": "70601561", "PTE. CJ 22OZ": "70601561",
        "PTE. HU 12OZ": "74621774", "PTE. LIGHT HU 12OZ": "74621774", "BRAHMA LIGHT HU 12OZ": "7468973200194",
        "BRAHMA LIGHT HU": "7468973200194", "BRAHMA LIGHT HU 16/650M": "7468973200200", "BRAHMA LIGHT 650ML": "7468973200200",
        "CORONA EXTRA 330ML": "7503034941200", "CORONA CERO 355ML": "750304423180", "MICHELOB ULTRA 355ML": "7422110104967",
        "THE ONE HU 12OZ": "74601325", "THE ONE HU 22OZ": "74601127", "CLAMATO COCTEL TOMATE C": "01484035",
        "ENRIQUILLO SODA 400 ML": "7463172803733", "GATORADE FRUIT PUNCH": "7460548000154", "GATORADE NARANJA": "92735",
        "GATORADE UVA": "92736", "FOUR LOKO MARACUYA": "849806004962", "FOUR LOKO PONCHE DE FRUTAS": "849806001220",
        "FOUR LOKO GREEN": "849806001855", "FOUR LOKO PURPLE": "849806002746", "FOUR LOKO GOLD": "849806001756",
        "FOUR LOKO SANDIA": "849806001206", "FOUR LOKO WHITE": "849806005754", "ALOE PURE PLUS ORIGINAL": "8809125063011",
        "MY COCO PURE PLUS": "8809125063011",
        # Álvarez & Sánchez (Pág. 3)[cite: 3]
        "SANTA HELENA MERLOT 75 CL": "7804300120986",[cite: 3]
        "SANTA HELENA RESERVADO RED BLEND 75 CL": "7804300150082",[cite: 3]
        "SANTA HELENA SAUVIGNON BLANC 75 CL": "7804300150041",[cite: 3]
        "SANTA HELENA VINO DULCE TINTO 75 CL": "7804300149307",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 1.5 LT": "8420976010063",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 375 CL": "8420976010087",[cite: 3]
        "SANTIAGO RUIZ ALBARIÑO 75 CL": "842097601070",[cite: 3]
        "SCHWEPPES AGUA TONICA 4 PACK 18 CL": "2000011980849",[cite: 3]
        "SCHWEPPES TONICA 1 LT": "2117974",[cite: 3]
        "SCHWEPPES TONICA ZERO 1 LT": "2138531",[cite: 3]
        "SELA BODEGAS RODA VINO TINTO 75 CL": "8014396003073",[cite: 3]
        "SOLAN DE CABRAS AGUA MINERAL NAT 1.5 LT": "8436538810767"[cite: 3]
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
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Memoria Independiente Activa</p>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: PROCESAR FACTURA
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador Inteligente por Proveedor</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Selecciona el proveedor exacto para aplicar su perfil de formato independiente.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""
    if "paginacion_detectada" not in st.session_state: st.session_state["paginacion_detectada"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    lista_proveedores = list(st.session_state["supplier_memory"].keys()) + ["➕ Registrar Nuevo Proveedor"]
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        prov_seleccionado = st.selectbox("🏢 Selecciona el Proveedor", lista_proveedores)
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu factura o tique (PDF multi-página o Imagen)", type=["pdf", "png", "jpg", "jpeg"])
    
    if archivo_subido is not None:
        if st.button("🚀 Procesar Factura con Perfil Aislado"):
            with st.spinner(f"🚀 Leyendo factura bajo la configuración exclusiva de '{prov_seleccionado}'..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave de API configurada.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    archivo_subido.seek(0)
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in f_type.lower() else Image.open(io.BytesIO(file_bytes))

                    supp_mem = st.session_state["supplier_memory"]
                    
                    if prov_seleccionado == "➕ Registrar Nuevo Proveedor":
                        instruccion_proveedor = "Analiza el documento, identifica el nombre comercial del proveedor y extrae todos sus renglones con descripción, cantidad, unidad y monto neto."
                        prov_nombre_objetivo = "NUEVO PROVEEDOR"
                    else:
                        instruccion_proveedor = supp_mem[prov_seleccionado].get("instruccion_prompt", "Extrae todos los ítems.")
                        prov_nombre_objetivo = prov_seleccionado

                    prompt_unificado = (
                        f"Estás procesando un documento del proveedor '{prov_nombre_objetivo}'. "
                        f"Instrucción de formato específica para este proveedor: {instruccion_proveedor} "
                        "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                        '{"paginacion": "1 de 1", "proveedor_detectado": "...", "subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"descripcion": "...", "tamano": "...", "cantidad": 1.0, "unidad": "...", "monto_neto": 0.0}]}. '
                        "Respuesta JSON pura."
                    )

                    archivo_subido.seek(0)
                    response = model.generate_content([image_input, prompt_unificado])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("```"): raw_text = raw_text[:-3]
                    
                    parsed_json = json.loads(raw_text.strip())
                    
                    prov_a_usar = prov_seleccionado
                    if prov_seleccionado == "➕ Registrar Nuevo Proveedor":
                        prov_a_usar = str(parsed_json.get("proveedor_detectado") or "PROVEEDOR NUEVO").upper().strip()
                        if prov_a_usar not in supp_mem:
                            supp_mem[prov_a_usar] = {
                                "nombre": prov_a_usar,
                                "tipo_formato": "personalizado",
                                "instruccion_prompt": instruccion_proveedor
                            }
                            st.session_state["supplier_memory"] = supp_mem
                            save_json_file(SUPPLIER_MEMORY_FILE, supp_mem)

                    st.session_state["factura_data"] = parsed_json
                    st.session_state["prov_activo"] = prov_a_usar
                    st.session_state["paginacion_detectada"] = str(parsed_json.get("paginacion", "1 de 1"))
                    
                    st.success(f"✅ ¡Factura procesada con el perfil de **{prov_a_usar}**! Se extrajeron **{len(parsed_json.get('items', []))}** renglones.")
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

            for idx, item in enumerate(items, start=1):
                raw_desc = item.get("descripcion", "")
                raw_tam = item.get("tamano", "")
                
                nombre_completo = limpiar_nombre_producto(raw_desc, raw_tam)
                m_med = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z))', str(raw_tam + " " + raw_desc).upper())
                presentacion_limpia = m_med.group(1) if m_med else (raw_tam if raw_tam else "S/P")

                codigo_final = buscar_en_catalogo_maestro(nombre_completo, presentacion_limpia)

                cant_compra = safe_float(item.get("cantidad"), 1.0)
                unidad = str(item.get("unidad", ""))
                monto_neto_fila = safe_float(item.get("monto_neto") or item.get("impuesto_neto"), 0.0)

                empaque = parse_empaque(raw_tam, unidad, raw_desc)
                total_unidades = int(cant_compra * empaque)
                costo_unitario_real = round(monto_neto_fila / total_unidades, 2) if total_unidades > 0 else 0.0

                if costo_unitario_real > 0:
                    precio_con_utilidad = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(precio_con_utilidad * 1.18)
                else:
                    precio_venta = 0.0

                preview_rows.append({
                    "No.": idx, "Producto": nombre_completo, "Código EAN Asignado": codigo_final,
                    "Cant. Compra": cant_compra, "Empaque": empaque, "Stock (Unidades)": total_unidades,
                    "Costo Unit. Real": costo_unitario_real, "Precio Venta": precio_venta
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
    st.markdown("<h2>📁 Catálogo Maestro de Productos</h2>", unsafe_allow_html=True)
    master_data = st.session_state.get("master_catalog", {})
    if master_data:
        df_show = pd.DataFrame([{"Producto": k, "Código EAN Oficial": v} for k, v in master_data.items()])
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
