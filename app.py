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

# Archivos persistentes de memoria, reglas y último maestro
MEMORY_FILE = "proveedores_memoria.json"
OVERRIDES_FILE = "mapeo_productos_overrides.json"
SAVED_MASTER_FILE = "ultimo_maestro_pos.xlsx"

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

# Carga automática del último maestro guardado en disco si existe
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
        st.sidebar.success(f"📂 Maestro anterior cargado automáticamente: {len(master_dict)} productos.")
        st.session_state["master_loaded_once"] = True
    except Exception as e:
        print(f"Error cargando maestro previo: {e}")

master_file_uploaded = st.sidebar.file_uploader("Actualizar Maestro (Sube nuevo archivo)", type=["xlsx", "xls", "csv"], key="master_inv_file")

if master_file_uploaded is not None:
    try:
        if master_file_uploaded.name.endswith('.csv'):
            df_master = pd.read_csv(master_file_uploaded)
            # Guardar copia en excel para persistencia uniforme
            df_master.to_excel(SAVED_MASTER_FILE, index=False)
        else:
            df_master = pd.read_excel(master_file_uploaded)
            # Guardar archivo subido directamente
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
                
        st.sidebar.success(f"✅ Nuevo maestro guardado y cargado: {len(master_dict)} productos.")
    except Exception as e:
        st.sidebar.error(f"Error al procesar el maestro: {e}")
elif os.path.exists(SAVED_MASTER_FILE) and not master_dict:
    # Cargar en memoria si el uploader está vacío pero el archivo existe en disco
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

# Equivalencias personalizables
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO",
        "DOBEL": "MAESTRO DOBEL"
    }

st.sidebar.markdown("---")
st.sidebar.title("🔄 Equivalencias y Sinónimos")
with st.sidebar.expander("Ver / Editar Equivalencias"):
    eq_key = st.text_input("Término del Proveedor")
    eq_val = st.text_input("Equivalente en tu POS")
    if st.button("➕ Agregar Regla"):
        if eq_key and eq_val:
            st.session_state["custom_equivalences"][eq_key.strip().upper()] = eq_val.strip().upper()
            st.success("¡Regla agregada!")
            st.rerun()
            
    if st.session_state["custom_equivalences"]:
        st.markdown("**Reglas activas:**")
        to_remove = []
        for k, v in st.session_state["custom_equivalences"].items():
            if st.checkbox(f"{k} ➔ {v}", value=True, key=f"eq_{k}") == False:
                to_remove.append(k)
        if to_remove:
            for r in to_remove:
                del st.session_state["custom_equivalences"][r]
            st.rerun()

# ==========================================
# MOTOR DE INTELIGENCIA Y RANGOS DE EMPAREJAMIENTO SEGURO
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
    for prov, pos in st.session_state["custom_equivalences"].items():
        if prov in upper_text:
            upper_text = upper_text.replace(prov, pos)
            
    for abbr, full in SYNONYMS_MAP.items():
        upper_text = upper_text.replace(abbr, full)
        
    cleaned = re.sub(r'[^A-Z0-9\s]', ' ', upper_text)
    stopwords = {"DE", "EL", "LA", "LOS", "LAS", "Y", "EN", "UN", "UNA", "CON", "CL", "ML", "L", "BCA", "BOT"}
    tokens = [t for t in cleaned.split() if t not in stopwords]
    return tokens, upper_text

def extract_volume_token(text):
    match = re.search(r'(\d+\s*(?:ML|L|CL))', str(text).upper())
    if match:
        v = match.group(1).replace(" ", "")
        if "L" in v and "ML" not in v:
            try:
                num = float(re.sub(r'[^0-9.]', '', v))
                return f"{int(num * 1000)}ML"
            except:
                pass
        return v
    return ""

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "Sin Maestro Cargado"
    
    clean_desc_key = item_description.strip().upper()
    
    # 1. Prioridad absoluta a reglas manuales o aprendidas previamente
    if clean_desc_key in st.session_state["product_overrides"]:
        return st.session_state["product_overrides"][clean_desc_key], "Actualizado (Regla Guardada)"

    prov_tokens, norm_prov = clean_and_normalize(item_description)
    prov_volume = extract_volume_token(item_description)
    
    # 2. Búsqueda exacta normalizada
    for m_name, m_code in master_dict.items():
        _, norm_m = clean_and_normalize(m_name)
        if m_name == norm_prov or m_name == clean_desc_key or norm_m == norm_prov:
            return m_code, "Actualizado (Exacto)"
            
    best_match_code = original_code
    best_match_name = ""
    max_matched_tiers = 0
    highest_score = 0.0

    prov_set = {t for t in prov_tokens if len(t) > 2}

    for m_name in master_names:
        m_tokens, _ = clean_and_normalize(m_name)
        m_volume = extract_volume_token(m_name)
        master_set = {t for t in m_tokens if len(t) > 2}

        if not prov_set or not master_set:
            continue

        matched_tiers = 0

        # RANGO 1: Coincidencia obligatoria de Marca o Término Principal
        common_tokens = prov_set.intersection(master_set)
        if not common_tokens:
            continue
        matched_tiers += 1

        # RANGO 2: Coincidencia de Volumen / Presentación si ambos están especificados
        if prov_volume and m_volume:
            if prov_volume != m_volume:
                continue
        matched_tiers += 1

        # RANGO 3: Coincidencia semántica de similitud alta
        union_tokens = prov_set.union(master_set)
        jaccard = len(common_tokens) / len(union_tokens)
        
        score = jaccard
        for t in common_tokens:
            if len(t) > 3:
                score += 0.25

        if matched_tiers >= 2 and score >= 0.50:
            if score > highest_score:
                highest_score = score
                max_matched_tiers = matched_tiers
                best_match_code = master_dict[m_name]
                best_match_name = m_name

    if highest_score >= 0.55 and max_matched_tiers >= 2:
        return best_match_code, f"Actualizado (IA por Rangos: {best_match_name})"

    return original_code, "⚠️ Conserva Código Original (Sin Rango Seguro)"

# Función centralizada con Memoria de Proveedores
def process_invoice_with_ai(file_obj, file_type):
    memory_context = ""
    known_mem = st.session_state["provider_memory"]
    if known_mem:
        memory_context = "MEMORIA HISTÓRICA DE FORMATOS DE PROVEEDORES:\n"
        for rnc, info in known_mem.items():
            memory_context += f"- Proveedor RNC {rnc} ({info.get('nombre', '')}): {info.get('nota_formato', 'Formato estándar de cajas con empaques fraccionados.')}\n"

    prompt_text = (
        f"{memory_context}\n"
        "Analiza esta factura detalladamente. Extrae los datos de cabecera con absoluta precisión: 'emisor_rnc', 'emisor_nombre', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (número de cajas compradas), 'empaque' (unidades individuales que trae la caja, interpretando formatos como 12/75CL -> 12, 6/4PACK -> 24 o 6, etc.), y 'costo_sin_itbis' (EL COSTO UNITARIO REAL POR CADA PIEZA INDIVIDUAL: toma el precio neto total de la línea y divídelo estrictamente entre cantidad * empaque). "
        "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre comercial del producto y su presentación o tamaño limpio (ej: 'MAESTRO DOBEL DIAMANTE 700 ML', 'EVIAN 75 CL', 'BLUE LABEL 750 ML'), eliminando códigos internos, diagonales de empaque y textos redundantes. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
        "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
    )

    parsed_data = None
    success_msg = ""
    
    active_key = paid_api_key if st.session_state["use_paid_now"] else free_key_1
    if not active_key and not st.session_state["use_paid_now"]:
        active_key = free_key_2

    try:
        genai.configure(api_key=active_key if active_key else paid_api_key)
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
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3]
        
        parsed_data = json.loads(raw_text.strip())
        success_msg = "✅ ¡Factura procesada con éxito y formato aprendido!"
        
        rnc_key = str(parsed_data.get("emisor_rnc", "")).strip()
        nombre_prov = str(parsed_data.get("emisor_nombre", "Proveedor Desconocido")).strip()
        
        if rnc_key and rnc_key not in st.session_state["provider_memory"]:
            st.session_state["provider_memory"][rnc_key] = {
                "nombre": nombre_prov if nombre_prov and nombre_prov != "None" else f"Proveedor RNC {rnc_key}",
                "nota_formato": "Formato de cajas con empaques y costos unitarios procesados exitosamente."
            }
            save_json_file(MEMORY_FILE, st.session_state["provider_memory"])

        if st.session_state["use_paid_now"]:
            success_msg = "✅ ¡Factura procesada usando la Versión de Pago!"
            
    except Exception as e:
        err_str = str(e)
        if ("429" in err_str or "Quota exceeded" in err_str) and not st.session_state["use_paid_now"]:
            try:
                genai.configure(api_key=free_key_2 if free_key_2 else paid_api_key)
                model2 = genai.GenerativeModel('gemini-3.6-flash')
                file_obj.seek(0)
                response = model2.generate_content([
                    {'mime_type': file_type, 'data': file_obj.read()},
                    prompt_text
                ])
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.endswith("
