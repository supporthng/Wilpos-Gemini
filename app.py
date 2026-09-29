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
    page_title="WilPOS - Sistema Multi-Proveedor Blindado", 
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
# GESTIÓN DE PERFILES Y CATÁLOGO BLINDADO
# ==========================================
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    default_profiles = {
        "CND / BEES": {
            "nombre": "CND / BEES",
            "tipo_formato": "tique_doble_linea_blindado",
            "instruccion_prompt": "Analiza este tique de CND / BEES donde cada ítem tiene dos líneas: la línea 1 con código, unidad (PC o UN) y descripción, y la línea 2 con cantidad, precio unitario (P.Unit) e impuesto neto. Extrae 'descripcion', 'tamano', 'cantidad', 'unidad' y 'precio_unitario'."
        },
        "ALVAREZ & SANCHEZ": {
            "nombre": "ALVAREZ & SANCHEZ",
            "tipo_formato": "factura_codigo_barras_impreso",
            "instruccion_prompt": "Analiza esta factura de ALVAREZ & SANCHEZ renglón por renglón. Extrae estrictamente la columna 'CODIGO' (código interno/SAP), la columna 'CODIGO DE BARRAS' (el código EAN impreso), la 'DESCRIPCION', el 'TAMAÑO', la cantidad y el 'VALOR'."
        },
        "GONZALEZ CUESTA & SUCS": {
            "nombre": "GONZALEZ CUESTA & SUCS",
            "tipo_formato": "factura_sap_desglose",
            "instruccion_prompt": "Analiza esta factura de GONZALEZ CUESTA & SUCS renglón por renglón. Extrae el código SAP, descripción, cantidad, unidad (UMV) y monto neto."
        }
    }
    for k, v in default_profiles.items():
        if k not in loaded_supps:
            loaded_supps[k] = v
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
        "ENRIQUILLO SODA 400 ML": "7463172803733", "GATORADE FRUIT PUNCH": "7460548000154", "GATORADE NARANJA": "052000324884",
        "GATORADE UVA": "052000324822", "FOUR LOKO MARACUYA": "849806004962", "FOUR LOKO PONCHE DE FRUTAS": "849806001220",
        "FOUR LOKO GREEN": "849806001855", "FOUR LOKO PURPLE": "849806002746", "FOUR LOKO GOLD": "849806001756",
        "FOUR LOKO SANDIA": "849806001206", "FOUR LOKO WHITE": "849806005754", 
        "ALOE PURE PLUS ORIGINAL 1.5 LT": "8809125063035",
        "ALOE PURE PLUS ORIGINAL": "8809125063011", 
        "MY COCO PURE PLUS": "8809125063011",
        "LICOR DE CAFE TIA MARIA 70 CL": "5012523233129",
        "TIA MARIA CAFE 70ML": "5012523233129",
        "VINO TINTO RESERVA CUNE 12/75 CL.": "8410591003045",
        "VINO TINTO MERLOT VIÑA TARAPACA 12/75 CL.": "7804304909934",
        "VINO TINTO RESERVA CAB SAUV TARAPACA 12/75 CL.": "7804304909039",
        "VINO TINTO RESERVA CARMENERE TARAPACA 12/75 CL.": "7804304902184",
        "VINO TINTO RESERVA MERLOT TARAPACA 12/75 CL.": "7804304909958",
        "VINO TINTO RED BLEND JUAN GIL(JUMILLA)22 12/75 CL.": "851115002706",
        "VINO TTIO ET.AMARILLA JUAN GIL(JUMILLA)23 12/75 CL.": "8437005068001",
        "VINO TTIO ETIO AZUL JUAN GIL (JUMILLA)22 6/75 CL.": "8437005068735",
        "VINO TTIO ETIO PLATA JUAN GIL(JUMILLA)22 12/75 CL.": "8437005068072",
        "VINO TTO CAB SAUV BOURBON RESERV JOSH 22 12/75 CL.": "857744011157",
        "VINO TTO CAB SAUV NORTH RESERVE JOSH 21 12/75 CL.": "031259004327",
        "VINO TTO SIX EIGHT NINE 689 12/75 CL.": "031259000046",
        "WHISKY ESCOCES MALTA 12 AÑOS GLEN GRANT 12/75 CL.": "051497455309",
        "VODKA INFUSIONS CITRUS SKYYY 12/75 CL.": "051497322618",
        "VODKA INFUSIONS RASPBERRY SKYYY 6/70 CL.": "8000040630269",
        "VODKA SKYY 75 CL.": "721059627504",
        "VODKA SKYY 75 CL": "721059637503",
        "VODKA SKYY 75 CL.": "721059007504"
    }
    for k, v in base_defaults.items():
        if k not in loaded_master: loaded_master[k] = v
    save_json_file(MASTER_CATALOG_FILE, loaded_master)
    st.session_state["master_catalog"] = loaded_master

if "codigos_manuales_sesion" not in st.session_state:
    st.session_state["codigos_manuales_sesion"] = {}

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
    if s_val == "088857003006": s_val = "088857003306"
    if 4 <= len(s_val) <= 14: return str(s_val)
    return "S/C"

def limpiar_nombre_producto(descripcion_raw, tamano_raw):
    desc = str(descripcion_raw or "").upper().strip()
    desc = re.sub(r'\s+\d{4,6}$', '', desc)
    desc = re.sub(r'\b(PC|UN|CAJA|CAJ|BOT|PZA|CAJ /)\b', '', desc)
    desc = re.sub(r'\s+', ' ', desc).strip()
    
    tam = str(tamano_raw or "").upper().strip()
    m_medida = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z|CL))', tam + " " + desc)
    medida_limpia = m_medida.group(1) if m_medida else ""
    
    if medida_limpia and medida_limpia not in desc:
        return f"{desc} {medida_limpia}".strip()
    
    return desc

def parse_empaque_proveedor(proveedor_nombre, tamano_txt="", unidad_txt="", descripcion_txt=""):
    combined = f"{str(tamano_txt or '')} {str(unidad_txt or '')} {str(descripcion_txt or '')}".upper()
    
    m_slash = re.search(r'\b(48|24|20|18|16|12|6|4)\s*/', combined)
    if m_slash: return int(m_slash.group(1))

    m_pza = re.search(r'\b(48|24|12|6|4)\s*PZA\b', combined)
    if m_pza: return int(m_pza.group(1))

    m_gen = re.search(r'\b(48|24|12|6|4)\b', combined)
    if m_gen: return int(m_gen.group(1))

    return 1

def buscar_en_catalogo_maestro(nombre_producto, presentacion=""):
    n_upper = str(nombre_producto or "").upper().strip()
    p_upper = str(presentacion or "").upper().strip()
    combined_query = f"{n_upper} {p_upper}".strip()

    manual_dict = st.session_state.get("codigos_manuales_sesion", {})
    if combined_query in manual_dict: return manual_dict[combined_query]
    if n_upper in manual_dict: return manual_dict[n_upper]

    master_dict = st.session_state.get("master_catalog", {})
    n_clean = re.sub(r'[^A-Z0-9]', '', combined_query)
    
    for m_key, m_code in master_dict.items():
        m_clean = re.sub(r'[^A-Z0-9]', '', str(m_key).upper())
        if n_clean == m_clean or n_clean in m_clean or m_clean in n_clean:
            return clean_ean_code(m_code)

    return "S/C"

st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS Multi-Proveedor</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='font-size: 0.8rem; color: #10b981; font-weight: 600;'>🟢 Edición Inline en Tabla Activa</p>", unsafe_allow_html=True)

# ==========================================
# MÓDULO 1: PROCESAR FACTURA
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador Inteligente Multi-Proveedor (Multi-Página)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tus páginas. Puedes editar cualquier código EAN directamente haciendo clic en la celda de la tabla inferior.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        st.info("💡 Sube tus archivos de factura (puedes seleccionar varias páginas simultáneamente).")
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivos_subidos = st.file_uploader("📂 Sube tus páginas (imágenes o PDFs)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if archivos_subidos:
        if st.button("🚀 Procesar Páginas y Consolidar Inventario"):
            with st.spinner("🔍 Analizando páginas y acumulando ítems..."):
                try:
                    if not gemini_key: raise ValueError("No hay clave de API configurada.")
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    
                    todos_los_items = []
                    subtotal_acum = 0.0
                    itbis_acum = 0.0
                    descuentos_acum = 0.0
                    total_acum = 0.0
                    prov_encontrado = "PROVEEDOR GENERAL"

                    for idx_f, archivo_subido in enumerate(archivos_subidos):
                        archivo_subido.seek(0)
                        file_bytes = archivo_subido.read()
                        f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                        image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in f_type.lower() else Image.open(io.BytesIO(file_bytes))

                        if idx_f == 0:
                            prompt_deteccion = (
                                "Analiza este documento comercial e identifica estrictamente el nombre comercial del proveedor emisor. "
                                "Devuelve únicamente un JSON: {\"proveedor_detectado\": \"NOMBRE DEL PROVEEDOR\"}"
                            )
                            response_det = model.generate_content([image_input, prompt_deteccion])
                            raw_det_text = response_det.text.strip()
                            if raw_det_text.startswith("```json"): raw_det_text = raw_det_text[7:]
                            if raw_det_text.endswith("```"): raw_det_text = raw_det_text[:-3]
                            det_json = json.loads(raw_det_text.strip())
                            prov_raw = str(det_json.get("proveedor_detectado", "PROVEEDOR GENERAL")).upper().strip()
                            
                            supp_mem = st.session_state["supplier_memory"]
                            for p_key in supp_mem.keys():
                                if p_key in prov_raw or prov_raw in p_key:
                                    prov_encontrado = p_key
                                    break
                            if prov_encontrado == "PROVEEDOR GENERAL":
                                prov_encontrado = prov_raw

                        instruccion_proveedor = st.session_state["supplier_memory"].get(prov_encontrado, {}).get("instruccion_prompt", "Extrae todos los ítems.")

                        prompt_unificado = (
                            f"Estás procesando la página {idx_f+1} de la factura del proveedor: '{prov_encontrado}'. "
                            f"Instrucción específica: {instruccion_proveedor} "
                            "Extrae 'codigo_barras', 'descripcion', 'tamano', 'cantidad', 'unidad' y 'valor' (o 'importe'). "
                            "Devuelve un JSON puro con esta estructura exacta y llaves en minúscula: "
                            '{"subtotal": 0.0, "itbis": 0.0, "descuentos": 0.0, "total": 0.0, "items": [{"codigo_barras": "...", "descripcion": "...", "tamano": "...", "cantidad": 1.0, "unidad": "CAJA", "precio_unitario": 0.0, "monto_neto": 0.0}]}. '
                            "Respuesta JSON pura."
                        )

                        archivo_subido.seek(0)
                        response = model.generate_content([image_input, prompt_unificado])
                        raw_text = response.text.strip()
                        if raw_text.startswith("```json"): raw_text = raw_text[7:]
                        if raw_text.endswith("```"): raw_text = raw_text[:-3]
                        
                        parsed_page = json.loads(raw_text.strip())
                        todos_los_items.extend(parsed_page.get("items", []))
                        subtotal_acum += safe_float(parsed_page.get("subtotal"))
                        itbis_acum += safe_float(parsed_page.get("itbis"))
                        descuentos_acum += safe_float(parsed_page.get("descuentos"))
                        total_acum += safe_float(parsed_page.get("total"))

                    factura_consolidada = {
                        "proveedor_detectado": prov_encontrado,
                        "subtotal": subtotal_acum,
                        "itbis": itbis_acum,
                        "descuentos": descuentos_acum,
                        "total": total_acum,
                        "items": todos_los_items
                    }

                    st.session_state["factura_data"] = factura_consolidada
                    st.session_state["prov_activo"] = prov_encontrado
                    
                    st.success(f"🎯 **¡Proceso exitoso!** Se procesaron {len(archivos_subidos)} página(s) de **{prov_encontrado}** con un total de **{len(todos_los_items)} renglones** consolidados.")
                except Exception as e:
                    st.error(f"⚠️ Error al procesar: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state["factura_data"] is not None:
        data_resp = st.session_state["factura_data"]
        items = data_resp.get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")
        
        subtotal_val = safe_float(data_resp.get("subtotal"))
        itbis_val = safe_float(data_resp.get("itbis"))
        total_descuentos = safe_float(data_resp.get("descuentos"))
        total_val = safe_float(data_resp.get("total"))

        total_importe_neto = sum(safe_float(i.get("monto_neto") or i.get("importe")) for i in items)
        if subtotal_val == 0.0 and items: subtotal_val = total_importe_neto
        if total_val == 0.0 and items: total_val = subtotal_val * 1.18

        st.markdown(f"### 📊 Dashboard Financiero Consolidado | Proveedor: {prov_actual}")
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1: st.metric(label="Subtotal / Bruto", value=f"${subtotal_val:,.2f}")
        with col_m2: st.metric(label="ITBIS Total", value=f"${itbis_val:,.2f}")
        with col_m3: st.metric(label="Descuentos", value=f"${total_descuentos:,.2f}")
        with col_m4: st.metric(label="Total General", value=f"${total_val:,.2f}")
            
        st.markdown("---")

        if items:
            st.markdown(f"### 📋 Detalle de Renglones Consolidados ({len(items)} ítems totales)")
            st.info("✏️ **Edición Directa:** Puedes hacer clic sobre cualquier celda en la columna **'Código EAN Asignado'** de la tabla para corregirlo al instante.")

            raw_preview_rows = []
            master_dict = st.session_state.get("master_catalog", {})
            manual_sesion = st.session_state.get("codigos_manuales_sesion", {})

            for idx, item in enumerate(items, start=1):
                raw_desc = item.get("descripcion", "")
                raw_tam = item.get("tamano", "")
                
                nombre_completo = limpiar_nombre_producto(raw_desc, raw_tam)
                cant_compra = safe_float(item.get("cantidad"), 1.0)
                unidad = str(item.get("unidad", ""))
                
                empaque = parse_empaque_proveedor(prov_actual, raw_tam, unidad, raw_desc)
                total_unidades = int(cant_compra * empaque)

                m_med = re.search(r'(\d+\s*(?:ML|L|LT|G|KG|OZ|Z|CL))', str((raw_tam or "") + " " + (raw_desc or "")).upper())
                presentacion_limpia = m_med.group(1) if m_med else (raw_tam if raw_tam else "S/P")
                nombre_display_excel = f"{nombre_completo} {presentacion_limpia}".strip()

                # Determinar código inicial
                if nombre_display_excel in manual_sesion:
                    codigo_final = manual_sesion[nombre_display_excel]
                else:
                    codigo_extraido = clean_ean_code(item.get("codigo_barras", ""))
                    if codigo_extraido == "S/C":
                        codigo_final = buscar_en_catalogo_maestro(nombre_display_excel, presentacion_limpia)
                    else:
                        codigo_final = codigo_extraido

                p_unit_extraido = safe_float(item.get("precio_unitario"), 0.0)
                monto_neto_linea = safe_float(
                    item.get("monto_neto") or 
                    item.get("importe") or 
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

                raw_preview_rows.append({
                    "No.": idx,
                    "Producto": nombre_display_excel,
                    "Código EAN Asignado": str(codigo_final),
                    "Cant. Compra": cant_compra,
                    "Empaque": empaque,
                    "Stock (Unidades)": total_unidades,
                    "Costo Unit. Real": costo_unitario_real,
                    "Precio Venta": precio_venta,
                    "_presentacion": presentacion_limpia
                })

            df_to_edit = pd.DataFrame(raw_preview_rows)
            
            # Tabla interactiva con celdas editables inline
            edited_df = st.data_editor(
                df_to_edit.drop(columns=["_presentacion"]),
                use_container_width=True,
                hide_index=True,
                key="grid_inventario_editable"
            )

            # Guardar automáticamente los cambios realizados en la tabla en el catálogo maestro y sesión
            for i, row in edited_df.iterrows():
                p_name = row["Producto"]
                nuevo_code_editado = clean_ean_code(row["Código EAN Asignado"])
                if nuevo_code_editado != "S/C":
                    st.session_state["codigos_manuales_sesion"][p_name] = nuevo_code_editado
                    master_dict[p_name] = nuevo_code_editado
            
            save_json_file(MASTER_CATALOG_FILE, master_dict)
            st.session_state["master_catalog"] = master_dict

            # Construir Excel con los datos editados
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            for idx_row, row in edited_df.iterrows():
                p_name = row["Producto"]
                p_code = row["Código EAN Asignado"]
                p_presentacion = raw_preview_rows[idx_row]["_presentacion"]
                ws.append([
                    p_name, p_presentacion, str(p_code), prov_actual, "producto",
                    row["Precio Venta"], row["Costo Unit. Real"], row["Stock (Unidades)"], 0.18, "unidad", row["Empaque"]
                ])

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            st.download_button(
                label=f"📥 Descargar Excel Consolidado - {prov_actual}",
                data=excel_buffer.getvalue(),
                file_name=f"Inventario_Master_{prov_actual.replace(' ', '_')}_Consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
                nuevo_prod_nombre = st.text_input("Nombre / Descripción del Producto")
            with col_m2:
                nuevo_prod_codigo = st.text_input("Código EAN / SAP Oficial")
            
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
        st.data_editor(df_show, use_container_width=True, hide_index=True, key="grid_maestro_editable")

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
