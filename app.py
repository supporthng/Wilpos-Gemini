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

# Archivo persistente de memoria de proveedores
MEMORY_FILE = "proveedores_memoria.json"

def load_provider_memory():
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_provider_memory(memory_dict):
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(memory_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error guardando memoria: {e}")

if "provider_memory" not in st.session_state:
    st.session_state["provider_memory"] = load_provider_memory()

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
st.sidebar.title("🧠 Memoria de Proveedores")
with st.sidebar.expander("Ver Proveedores Aprendidos"):
    mem = st.session_state["provider_memory"]
    if mem:
        st.write(f"Total proveedores en memoria: {len(mem)}")
        for rnc, info in mem.items():
            st.markdown(f"- **RNC:** `{rnc}` ({info.get('nombre', 'Sin nombre')})")
        if st.button("🗑️ Limpiar Memoria de Proveedores"):
            st.session_state["provider_memory"] = {}
            if os.path.exists(MEMORY_FILE):
                os.remove(MEMORY_FILE)
            st.success("¡Memoria reseteada!")
            st.rerun()
    else:
      st.info("Aún no hay proveedores aprendidos. Se registrarán automáticamente al procesar facturas.")

st.sidebar.markdown("---")
st.sidebar.title("🗂️ Maestro de Inventario POS")
master_file_uploaded = st.sidebar.file_uploader("Sube tu archivo Maestro", type=["xlsx", "xls", "csv"], key="master_inv_file")

master_dict = {}
master_names = []
if master_file_uploaded is not None:
    try:
        if master_file_uploaded.name.endswith('.csv'):
            df_master = pd.read_csv(master_file_uploaded)
        else:
            df_master = pd.read_excel(master_file_uploaded)
        
        cols = [c.lower() for c in df_master.columns]
        name_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'nombre' in c or 'descripcion' in c), df_master.columns[0])
        code_col = next((df_master.columns[i] for i, c in enumerate(cols) if 'codigo' in c or 'barra' in c or 'barcode' in c), df_master.columns[1])
        
        for _, row in df_master.iterrows():
            p_name = str(row[name_col]).strip().upper()
            p_code = str(row[code_col]).strip()
            if p_name and p_name != "NAN":
                master_dict[p_name] = p_code
                master_names.append(p_name)
                
        st.sidebar.success(f"✅ Maestro cargado: {len(master_dict)} productos.")
    except Exception as e:
        st.sidebar.error(f"Error al leer el maestro: {e}")

# Diccionario de equivalencias personalizables guardado en session_state
if "custom_equivalences" not in st.session_state:
    st.session_state["custom_equivalences"] = {
        "BARCELO 40 ANIVERSARIO": "IMPERIAL PREMIUM BLEND 40 AÑOS",
        "BARCELO IMPERIAL PORTO": "IMPERIAL PORTO"
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
# MOTOR DE NORMALIZACIÓN Y COINCIDENCIA AVANZADA
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
    "BOTELLA ": "",
    "BOTELL ": ""
}

def clean_and_tokenize(text):
    upper_text = str(text).upper()
    for prov, pos in st.session_state["custom_equivalences"].items():
        if prov in upper_text:
            upper_text = upper_text.replace(prov, pos)
            
    for abbr, full in SYNONYMS_MAP.items():
        upper_text = upper_text.replace(abbr, full)
        
    cleaned = re.sub(r'[^A-Z0-9\s]', ' ', upper_text)
    stopwords = {"DE", "EL", "LA", "LOS", "LAS", "Y", "EN", "UN", "UNA", "CON"}
    tokens = [t for t in cleaned.split() if t not in stopwords]
    return set(tokens), upper_text

def validate_with_master(item_description, original_code):
    if not master_dict:
        return original_code, "Sin Maestro Cargado"
    
    desc_tokens, norm_desc = clean_and_tokenize(item_description)
    
    for m_name, m_code in master_dict.items():
        if m_name == norm_desc or m_name == item_description.strip().upper():
            return m_code, "Actualizado (Exacto)"
            
    best_match_code = original_code
    best_match_name = ""
    highest_score = 0.0
    
    for m_name in master_names:
        m_tokens, _ = clean_and_tokenize(m_name)
        if not desc_tokens or not m_tokens:
            continue
            
        intersection = desc_tokens.intersection(m_tokens)
        union = desc_tokens.union(m_tokens)
        jaccard_score = len(intersection) / len(union)
        
        weight = 1.0
        for token in intersection:
            if token.isdigit() or len(token) > 3:
                weight += 0.25
                
        final_score = jaccard_score * weight
        
        if final_score > highest_score and len(intersection) >= 2:
            highest_score = final_score
            best_match_code = master_dict[m_name]
            best_match_name = m_name
            
    if highest_score >= 0.45:
        return best_match_code, f"Actualizado (IA Semántica: {best_match_name})"
        
    matches = difflib.get_close_matches(norm_desc, master_names, n=1, cutoff=0.45)
    if matches:
        matched_name = matches[0]
        return master_dict[matched_name], f"Actualizado (Similitud: {matched_name})"

    return original_code, "⚠️ No Encontrado en Maestro"

# Función centralizada con Memoria Inteligente de Proveedores
def process_invoice_with_ai(file_obj, file_type):
    # Revisar si tenemos memoria previa de proveedores para inyectarla como contexto y asegurar cero errores
    memory_context = ""
    known_mem = st.session_state["provider_memory"]
    if known_mem:
        memory_context = "MEMORIA HISTÓRICA DE FORMATOS DE PROVEEDORES:\n"
        for rnc, info in known_mem.items():
            memory_context += f"- Proveedor RNC {rnc} ({info.get('nombre', '')}): {info.get('nota_formato', 'Formato estándar de cajas con empaques fraccionados.')}\n"

    prompt_text = (
        f"{memory_context}\n"
        "Analiza esta factura detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'subtotal', 'itbis', 'total'. "
        "Para cada ítem, extrae con absoluta precisión: 'codigo', 'descripcion', 'cantidad' (número de cajas compradas), 'empaque' (unidades individuales que trae la caja, interpretando formatos como 12/75CL -> 12, 6/4PACK -> 24 o 6, etc.), y 'costo_sin_itbis' (EL COSTO UNITARIO REAL POR CADA PIEZA INDIVIDUAL: toma el precio neto total de la línea y divídelo estrictamente entre cantidad * empaque). "
        "REGLA ESTRICTA PARA LA DESCRIPCIÓN: Limpia el texto de cada producto para incluir ÚNICAMENTE el nombre comercial del producto y su presentación o tamaño limpio (ej: 'EVIAN 75 CL', 'GINGER BEER SPICY 207 ML', 'BLUE LABEL 750 ML'), eliminando códigos internos, diagonales de empaque y textos redundantes. "
        "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
        '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "subtotal": 0.0, "itbis": 0.0, "total": 0.0, "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
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
        
        # APRENDIZAJE AUTOMÁTICO DE PROVEEDORES
        rnc_key = str(parsed_data.get("emisor_rnc", "")).strip()
        if rnc_key and rnc_key not in st.session_state["provider_memory"]:
            st.session_state["provider_memory"][rnc_key] = {
                "nombre": f"Proveedor RNC {rnc_key}",
                "nota_formato": "Formato de cajas con empaques y costos unitarios procesados exitosamente."
            }
            save_provider_memory(st.session_state["provider_memory"])

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
