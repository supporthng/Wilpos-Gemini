import io
import json
import os
import hashlib
import re
import difflib
import google.generativeai as genai
from PIL import Image
import streamlit as st
import openpyxl
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="WilPOS - Automatizador de Facturas", page_icon="📊", layout="wide")

# Configuración de Claves API desde secrets de Streamlit o variables de entorno
free_key_1 = st.secrets.get("GEMINI_API_KEY_1", os.environ.get("GEMINI_API_KEY_1", ""))
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", os.environ.get("GEMINI_API_KEY_2", ""))
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", os.environ.get("GEMINI_API_KEY_PAID", ""))

# Archivos persistentes de memoria y reglas
MEMORY_FILE = "proveedores_memoria.json"
OVERRIDES_FILE = "mapeo_productos_overrides.json"
SAVED_MASTER_FILE = "Inventario_Completo_2026-09-11.xlsx"
TEMPLATE_FILE = "Plantilla_Inventario_WilPOS.xlsx"

def load_json_file(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_json_file(filepath, data_dict):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando {filepath}: {e}")

if "provider_memory" not in st.session_state:
    st.session_state["provider_memory"] = load_json_file(MEMORY_FILE)

if "product_overrides" not in st.session_state:
    st.session_state["product_overrides"] = load_json_file(OVERRIDES_FILE)

# Funciones auxiliares de cálculo y formato
def safe_float(val, default=0.0):
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
        return float(val)
    except (ValueError, TypeError):
        return default

def safe_int(val, default=1):
    try:
        return int(val)
    except (ValueError, TypeError):
        return default

def round_to_nearest_5(x):
    return round(round(x / 5) * 5, 2)

if "quota_exceeded" not in st.session_state:
    st.session_state["quota_exceeded"] = False

if "use_paid_now" not in st.session_state:
    st.session_state["use_paid_now"] = False

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.title("Menú de Navegación")
modulo = st.sidebar.radio(
    "Selecciona el Módulo",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

st.sidebar.markdown("---")
st.sidebar.title("🧠 Memoria y Reglas POS")
with st.sidebar.expander("Ver Correcciones de Productos"):
    overrides = st.session_state["product_overrides"]
    if overrides:
        st.write(f"Total reglas de mapeo: {len(overrides)}")
        for prov_desc, pos_code in overrides.items():
            st.markdown(f"- `{prov_desc}` ➔ Código: **{pos_code}**")
        if st.button("🗑️ Limpiar Reglas de Mapeo"):
            st.session_state["product_overrides"] = {}
            if os.path.exists(OVERRIDES_FILE):
                os.remove(OVERRIDES_FILE)
            st.success("¡Reglas reseteadas!")
            st.rerun()
    else:
        st.info("No hay reglas manuales registradas.")

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")

master_dict = {}
master_names = []

if os.path.exists(SAVED_MASTER_FILE) and "master_loaded_once" not in st.session_state:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE)
        cols_s = [c.lower() for c in df_saved.columns]
        s_name_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'nombre' in c or 'descripcion' in c), df_saved.columns[0])
        s_code_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_saved.columns[1])
        
        for _, row in df_saved.iterrows():
            p_name = str(row[s_name_col]).strip().upper()
            p_code = str(row[s_code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
        st.sidebar.success(f"📂 Maestro cargado: {len(master_dict)} productos.")
        st.session_state["master_loaded_once"] = True
    except Exception as e:
        print(f"Error cargando maestro previo: {e}")

master_file_uploaded = st.sidebar.file_uploader("Actualizar Maestro (Sube nuevo archivo)", type=["xlsx", "xls", "csv"], key="master_inv_file")

if master_file_uploaded is not None:
    st.sidebar.warning("⚠️ Has subido un nuevo archivo maestro. Se requiere tu autorización para aplicarlo.")
    if st.sidebar.button("🔒 Autorizar y Guardar Nuevo Maestro"):
        try:
            if master_file_uploaded.name.endswith('.csv'):
                df_master = pd.read_csv(master_file_uploaded)
                df_master.to_excel(SAVED_MASTER_FILE, index=False)
            else:
                df_master = pd.read_excel(master_file_uploaded)
                with open(SAVED_MASTER_FILE, "wb") as f:
                    f.write(master_file_uploaded.getbuffer())
            
            master_dict = {}
            master_names = []
            cols = [c.lower() for c in df_master.columns]
            name_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_master.columns[0])
            code_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_master.columns[1])
            
            for _, row in df_master.iterrows():
                p_name = str(row[name_col]).strip().upper()
                p_code = str(row[code_col]).strip()
                if p_name and p_name != "NAN":
                    master_dict[p_name] = p_code
                    master_names.append(p_name)
                    
            st.sidebar.success(f"✅ Nuevo maestro autorizado y guardado: {len(master_dict)} productos.")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al procesar el maestro: {e}")
elif os.path.exists(SAVED_MASTER_FILE) and not master_dict:
    try:
        df_saved = pd.read_excel(SAVED_MASTER_FILE)
        cols_s = [c.lower() for c in df_saved.columns]
        s_name_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'nombre' in c or 'descripcion' in c), df_saved.columns[0])
        s_code_col = next((df_saved.columns[i] for i, c in enumerate(cols_s) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_saved.columns[1])
        
        for _, row in df_saved.iterrows():
            p_name = str(row[s_name_col]).strip().upper()
            p_code = str(row[s_code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
    except Exception:
        pass

st.sidebar.markdown("---")
st.sidebar.title("📄 Plantilla Oficial WilPOS")
template_uploaded = st.sidebar.file_uploader("Actualizar Plantilla Base (.xlsx)", type=["xlsx"], key="template_file_uploader")
if template_uploaded is not None:
    st.sidebar.warning("⚠️ Se ha cargado una nueva plantilla base. Requiere autorización.")
    if st.sidebar.button("🔒 Autorizar y Guardar Plantilla"):
        with open(TEMPLATE_FILE, "wb") as f:
            f.write(template_uploaded.getbuffer())
        st.sidebar.success("✅ Plantilla oficial autorizada y guardada.")
        st.rerun()

# Equivalencias Directas Completas y Definitivas
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL",
        "CLAMATO COCTEL TOMATE": "CLAMATO 221ML",
        "MY COCO PURE PLUS": "ALOE PURE PLUS COCONUT",
        "CORONA EXTRA": "CORONA PEQ 12 OZ",
        "CORONA CERO": "CERVEZA CORONA CERO 355ML",
        "THE ONE 12OZ": "THE ONE 355ml",
        "THE ONE 22OZ": "THE ONE 355ml",
        "THE ONE HU": "THE ONE 355ml",
        "PTE. HU": "PRESIDENTE PEQ. 12oz Regular",
        "PTE. CJ": "PRESIDENTE REG. 22  REGULAR oz",
        "PRESIDENTE 12 OZ": "PRESIDENTE PEQ. 12oz Regular",
        "PRESIDENTE 22 OZ": "PRESIDENTE REG. 22  REGULAR oz",
        "ENRIQUILLO SODA": "SODA, ENRRIQUILLO 400ML"
    }

# ==========================================
# MOTOR DE EMPAREJAMIENTO MAESTRO BLINDADO
# ==========================================
SYNONYMS_MAP = {
    "JW ": "JOHNNIE WALKER ",
    "JW.": "JOHNNIE WALKER",
    "BUCH ": "BUCHANANS ",
    "BUCHANAN": "BUCHANANS",
    "0.5 LT": "500ML",
    "0.5LT": "500ML",
    "1 LT": "1000ML",
    "1LT": "1000ML",
    "RON ": "",
    "TEQ ": "TEQUILA ",
    "BOTELLA ": "",
    "BOTELL ": ""
}

def clean_and_normalize(text):
    upper_text = str(text).upper()
    upper_text = re.sub(r'^\d{4,6}', '', upper_text).strip()

    for prov, pos in st.session_state["custom_equivalences"].items():
        if prov in upper_text:
            upper_text = upper_text.replace(prov, pos)
            
    for abbr, full in SYNONYMS_MAP.items():
        upper_text = upper_text.replace(abbr, full)
        
    cleaned = re.sub(r'[^A-Z0-9\s]', ' ', upper_text)
    stopwords = {"DE", "EL", "LA", "LOS", "LAS", "Y", "EN", "UN", "UNA", "CON", "CL", "ML", "L", "BCA", "BOT", "HU", "CJ", "LP", "OZ", "P", "C", "EX"}
    tokens = [t for t in cleaned.split() if t not in stopwords]
    return tokens, upper_text

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "Sin Maestro Cargado"
    
    clean_desc_key = item_description.strip().upper()
    clean_desc_key = re.sub(r'^\d{4,6}', '', clean_desc_key).strip()
    
    if clean_desc_key in st.session_state["product_overrides"]:
        return st.session_state["product_overrides"][clean_desc_key], "Actualizado (Regla Guardada)"

    for k, v in st.session_state["custom_equivalences"].items():
        if k in clean_desc_key:
            target_name = v.strip().upper()
            for m_name, m_code in master_dict.items():
                if m_name.strip().upper() == target_name:
                    return m_code, "Actualizado (Equivalencia Directa)"

    prov_tokens, norm_prov = clean_and_normalize(item_description)
    
    brand_keywords = ["BRAHMA", "HEINEKEN", "CORONA", "PRESIDENTE", "MICHELOB", "COORS", "MILLER", "THE ONE", "CLAMATO", "ALOE", "SODA"]
    prov_brand = next((b for b in brand_keywords if b in norm_prov), None)

    for m_name, m_code in master_dict.items():
        _, norm_m = clean_and_normalize(m_name)
        
        if prov_brand and prov_brand not in norm_m:
            continue

        if m_name == norm_prov or m_name == clean_desc_key or norm_m == norm_prov:
            return m_code, "Actualizado (Exacto)"
            
    best_match_code = original_code
    highest_score = 0.0

    prov_set = {t for t in prov_tokens if len(t) > 2 and not t.isdigit()}

    for m_name in master_names:
        m_tokens, _ = clean_and_normalize(m_name)
        master_set = {t for t in m_tokens if len(t) > 2 and not t.isdigit()}

        if not prov_set or not master_set:
            continue

        m_brand = next((b for b in brand_keywords if b in m_name.upper()), None)
        if prov_brand and m_brand and prov_brand != m_brand:
            continue

        common_keywords = prov_set.intersection(master_set)
        if not common_keywords:
            continue

        union = prov_set.union(master_set)
        score = len(common_keywords) / len(union)

        if score > highest_score and score >= 0.35:
            highest_score = score
            best_match_code = master_dict[m_name]

    if highest_score >= 0.35:
        return best_match_code, "Actualizado (IA Maestro Seguro)"

    return original_code, "⚠️ Conserva Código Original (Sin Match Seguro)"

def process_invoice_with_ai(file_obj, file_type):
    active_key = paid_api_key if st.session_state["use_paid_now"] else free_key_1
    if not active_key and not st.session_state["use_paid_now"]:
        active_key = free_key_2

    if not active_key and not paid_api_key:
        st.error("❌ No se encontró ninguna API Key de Gemini configurada en los secrets de Streamlit.")
        return None, ""

    memory_context = ""
    known_mem = st.session_state["provider_memory"]
    if known_mem:
        memory_context = "MEMORIA HISTÓRICA DE FORMATOS DE PROVEEDORES:\n"
        for rnc, info in known_mem.items():
            memory_context += f"- Proveedor RNC {rnc} ({info.get('nombre', '')}): {info.get('nota_formato', 'Formato estándar de cajas con empaques fraccionados.')}\n"

    prompt_text = (
        f"{memory_context}\n"
        "Analiza esta factura detalladamente. Extrae los datos de cabecera con absoluta precisión: 'emisor_rnc', 'emisor_nombre', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae con absoluta precisión: 'codigo', 'descripcion', 'cantidad' (número de bultos/paquetes comprados), 'empaque' (unidades individuales que trae el paquete. REGLA ESPECIAL PARA ALOE PURE PLUS: si el código es 92713, el empaque es estrictamente 10 unidades por paquete. Para otros productos usa su empaque real como 24, 16, 6, o 1 si es unitario), y 'costo_sin_itbis'. "
        "REGLA CRÍTICA ABSOLUTA PARA EL COSTO UNITARIO: El costo devuelto en 'costo_sin_itbis' DEBE SER OBLIGATORIAMENTE EL COSTO POR UNIDAD SUELTA (PIEZA INDIVIDUAL), NUNCA EL COSTO DEL PAQUETE COMPLETO. "
        "Para calcularlo correctamente: toma el 'Imp. Neto' total de la línea y divídelo estrictamente entre (Cantidad de Paquetes × Empaque). Es decir: costo_sin_itbis = Imp. Neto / (Cantidad * Empaque). "
        "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre comercial y presentación limpia, eliminando códigos de proveedor y textos redundantes. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
    )

    parsed_data = None
    success_msg = ""
    
    try:
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-3.6-flash')
        
        file_obj.seek(0)
        file_bytes = file_obj.read()
        
        response = model.generate_content([
            {'mime_type': file_type, 'data': file_bytes},
            prompt_text
        ])
        
        raw_text = response.text.strip()
        if raw_text.startswith("```json"):
            raw_text = raw_text[7:]
        if raw_text.endswith("
