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

# Configuración de Clave API
gemini_key = st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else os.environ.get("GEMINI_API_KEY")
if gemini_key:
    genai.configure(api_key=gemini_key)

# Archivos de Persistencia
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

# Estados de Sesión
if "supplier_memory" not in st.session_state:
    loaded_supps = load_json_file(SUPPLIER_MEMORY_FILE, "dict")
    if not loaded_supps:
        loaded_supps = {
            "ALVAREZ & SANCHEZ": {"nombre": "ALVAREZ & SANCHEZ", "notas_formato": "Desglose estándar."},
            "GONZALEZ CUESTA": {"nombre": "GONZALEZ CUESTA", "notas_formato": "Corporativo."}
        }
        save_json_file(SUPPLIER_MEMORY_FILE, loaded_supps)
    st.session_state["supplier_memory"] = loaded_supps

if "master_catalog" not in st.session_state:
    st.session_state["master_catalog"] = load_json_file(MASTER_CATALOG_FILE, "dict")

# ------------------------------------------
# FUNCIONES DE APOYO Y CRUCE
# ------------------------------------------
def safe_float(val, default=0.0):
    try: return float(val)
    except (ValueError, TypeError): return default

def round_to_nearest_5(x): 
    return float(round(round(x / 5) * 5))

def clean_ean_code(code_val):
    if not code_val: return "S/C"
    s_val = str(code_val).strip()
    if s_val.endswith('.0'): s_val = s_val[:-2]
    if s_val.lower() in ["nan", "none", "", "s/c", "sin codigo"] or len(s_val) < 7:
        return "S/C"
    return str(s_val)

def parse_empaque(unidad_txt="", descripcion_txt=""):
    combined = f"{str(unidad_txt)} {str(descripcion_txt)}".upper()
    m_caja = re.search(r'(?:CAJA|CAJ|PAQ|PACK|BLISTER)[^\d]*(\d+)', combined)
    if m_caja:
        val = int(m_caja.group(1))
        if 1 < val <= 120: return val

    m_slash = re.search(r'\b(48|24|16|12|6|10|20|30)\s*/', combined)
    if m_slash: return int(m_slash.group(1))

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

def llamada_segura_gemini(model, contents, max_intentos=3):
    """Realiza llamadas a Gemini manejando automáticamente el límite de cuota (Error 429) con espera inteligente."""
    for intento in range(max_intentos):
        try:
            return model.generate_content(contents)
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "quota" in err_str.lower():
                if intento < max_intentos - 1:
                    tiempo_espera = 25 * (intento + 1) # Espera progresiva de 25s, 50s...
                    st.warning(f"⚠️ Límite de cuota gratuito alcanzado (429). Pausando automáticamente por {tiempo_espera}s antes del reintento ({intento+1}/{max_intentos})...")
                    time.sleep(tiempo_espera)
                    continue
            raise e
    raise Exception("Se agotaron los reintentos automáticos por límite de cuota.")

# ==========================================
# MENÚ LATERAL DE NAVEGACIÓN
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7;'>⚡ WilPOS System</h3>", unsafe_allow_html=True)
menu_opcion = st.sidebar.radio("Navegación", ["📄 Procesar Factura", "📁 Catálogo Maestro EAN", "🏢 Gestionar Proveedores"])

# ==========================================
# MÓDULO 1: PROCESAR FACTURA (MULTI-PÁGINA)
# ==========================================
if menu_opcion == "📄 Procesar Factura":
    st.markdown("<h2>📄 Procesador Inteligente de Facturas (Multi-Página)</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube facturas de una o varias páginas. El sistema leerá la totalidad de las hojas y cruzará con tu Catálogo Maestro.</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "factura_data" not in st.session_state: st.session_state["factura_data"] = None
    if "prov_activo" not in st.session_state: st.session_state["prov_activo"] = ""

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    lista_proveedores = ["🔍 Detección Automática (Nuevo Proveedor)"] + list(st.session_state["supplier_memory"].keys())
    
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        prov_seleccionado = st.selectbox("🏢 Selecciona el Proveedor", lista_proveedores)
    with col_s2:
        margen_utilidad = st.number_input("⚙️ Margen Utilidad (%)", min_value=0.0, max_value=500.0, value=25.0, step=1.0)
        
    archivo_subido = st.file_uploader("📂 Sube tu factura (PDF multi-página o Imagen)", type=["pdf", "png", "jpg", "jpeg"])
    
    if archivo_subido is not None:
        if st.button("🚀 Procesar Todas las Páginas"):
            with st.spinner("Analizando documento completo (con control automático de reintentos por cuota)..."):
                try:
                    model = genai.GenerativeModel('gemini-3.8-flash')
                    file_bytes = archivo_subido.read()
                    f_type = getattr(archivo_subido, 'type', 'image/jpeg')
                    
                    image_input = {"mime_type": "application/pdf", "data": file_bytes} if "pdf" in f_type.lower() else Image.open(io.BytesIO(file_bytes))

                    prov_a_usar = prov_seleccionado
                    if prov_seleccionado == "🔍 Detección Automática (Nuevo Proveedor)":
                        prompt_det = "Identifica el nombre comercial del proveedor emisor en este documento. Devuelve un JSON puro: {'proveedor': 'NOMBRE'}"
                        resp_det = llamada_segura_gemini(model, [image_input, prompt_det])
                        txt_det = resp_det.text.strip()
                        if txt_det.startswith("```json"): txt_det = txt_det[7:]
                        if txt_det.endswith("```"): txt_det = txt_det[:-3]
                        prov_a_usar = str(json.loads(txt_det.strip()).get("proveedor") or "PROVEEDOR NUEVO").upper().strip()

                        supps = st.session_state["supplier_memory"]
                        if prov_a_usar not in supps:
                            supps[prov_a_usar] = {"nombre": prov_a_usar, "notas_formato": "Auto-registrado multi-página."}
                            st.session_state["supplier_memory"] = supps
                            save_json_file(SUPPLIER_MEMORY_FILE, supps)

                    prompt_main = (
                        f"Analiza este documento completo de compra (que puede contener una o varias páginas) del proveedor '{prov_a_usar}'. "
                        "Revisa todas las páginas de principio a fin y extrae ABSOLUTAMENTE TODOS LOS RENGLONES Y PRODUCTOS de todas las páginas, sin omitir ninguno. "
                        "Para cada renglón extrae en un JSON bajo la clave 'items': "
                        "1. 'descripcion': nombre del producto. "
                        "2. 'tamano': presentación (ej: '750 ML'). "
                        "3. 'codigo_factura': código que trae la factura (si no es un código de barras estándar válido, déjalo vacío). "
                        "4. 'cantidad': cantidad comprada (ej: 2.0). "
                        "5. 'unidad': unidad de empaque (ej: 'CAJA 12'). "
                        "6. 'valor_con_itbis': monto TOTAL INCLUYENDO ITBIS de esa línea. "
                        "Estructura JSON exacta: "
                        '{"items": [{"descripcion": "...", "tamano": "...", "codigo_factura": "...", "cantidad": 1.0, "unidad": "...", "valor_con_itbis": 0.0}]}. '
                        "Respuesta JSON pura."
                    )

                    response = llamada_segura_gemini(model, [image_input, prompt_main])
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"): raw_text = raw_text[7:]
                    if raw_text.endswith("```"): raw_text = raw_text[:-3]
                    
                    st.session_state["factura_data"] = json.loads(raw_text.strip())
                    st.session_state["prov_activo"] = prov_a_usar
                    st.success(f"✅ Documento procesado con éxito. Proveedor: **{prov_a_usar}**")
                except Exception as e:
                    st.error(f"⚠️ Error al procesar multi-página: {str(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

    # Vista previa y cruce maestro
    if st.session_state["factura_data"] is not None:
        items = st.session_state["factura_data"].get("items", [])
        prov_actual = st.session_state.get("prov_activo", "GENERAL")
        
        if items:
            st.markdown(f"### 📊 Total de ítems extraídos ({len(items)} renglones) | Proveedor: {prov_actual}")
            preview_rows = []
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.append(['Nombre', 'Presentación', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'ITBIS', 'Unidad Medida', 'Cantidad Empaque'])

            for idx, item in enumerate(items, start=1):
                desc = str(item.get("descripcion", "")).strip().upper()
                tamano = str(item.get("tamano", "")).strip().upper()
                nombre_completo = f"{desc} {tamano}".strip()
                
                cod_factura_raw = clean_ean_code(item.get("codigo_factura"))
                if cod_factura_raw != "S/C":
                    codigo_final = cod_factura_raw 
                else:
                    codigo_final = buscar_en_catalogo_maestro(nombre_completo) 

                cant_compra = safe_float(item.get("cantidad"), 1.0)
                unidad = str(item.get("unidad", ""))
                val_con_itbis = safe_float(item.get("valor_con_itbis"), 0.0)

                empaque = parse_empaque(unidad, desc)
                total_unidades = int(cant_compra * empaque)
                
                costo_sin_itbis_total = val_con_itbis / 1.18 if val_con_itbis > 0 else 0.0
                costo_unitario_real = round(costo_sin_itbis_total / total_unidades, 2) if total_unidades > 0 else 0.0

                if costo_unitario_real > 0:
                    precio_con_utilidad = costo_unitario_real * (1 + (margen_utilidad / 100.0))
                    precio_venta = round_to_nearest_5(precio_con_utilidad * 1.18)
                else:
                    precio_venta = 0.0

                preview_rows.append({
                    "No.": idx, "Producto": nombre_completo, "Código EAN Asignado": codigo_final,
                    "Costo Unit. Sin ITBIS": costo_unitario_real, "Precio Venta": precio_venta
                })

                ws.append([
                    nombre_completo, tamano if tamano else "S/P", str(codigo_final), prov_actual, "producto",
                    precio_venta, costo_unitario_real, total_unidades, 0.18, "unidad", empaque
                ])

            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            st.download_button(
                label=f"📥 Descargar Excel Importable - {prov_actual}",
                data=excel_buffer.getvalue(),
                file_name=f"Inventario_Master_{prov_actual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==========================================
# MÓDULO 2: CATÁLOGO MAESTRO EAN
# ==========================================
elif menu_opcion == "📁 Catálogo Maestro EAN":
    st.markdown("<h2>📁 Actualizar Catálogo Maestro de Productos</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b;'>Sube tu archivo de Excel con el catálogo oficial.</p>", unsafe_allow_html=True)
    st.markdown("---")

    master_file = st.file_uploader("📂 Sube tu Catálogo Maestro (Excel)", type=["xlsx"])
    if master_file is not None:
        df_master = pd.read_excel(master_file, dtype=str)
        cols = df_master.columns.tolist()
        
        col_c1, col_c2 = st.columns(2)
        with col_c1: col_name = st.selectbox("Columna con Nombre del Producto", cols)
        with col_c2: col_code = st.selectbox("Columna con Código de Barra EAN", cols)
        
        if st.button("🔄 Guardar Catálogo en Memoria"):
            temp_dict = {}
            count = 0
            for _, row in df_master.iterrows():
                p_name = str(row[col_name]).strip().upper()
                p_code = clean_ean_code(row[col_code])
                if p_name and p_code != "S/C":
                    temp_dict[p_name] = p_code
                    count += 1
            
            st.session_state["master_catalog"] = temp_dict
            save_json_file(MASTER_CATALOG_FILE, temp_dict)
            st.success(f"¡Catálogo maestro actualizado con éxito! Se cargaron **{count}** productos.")

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
