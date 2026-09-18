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
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS CSS
# ==========================================
st.set_page_config(
    page_title="WilPOS - Consulta Web de Productos", 
    page_icon="🌐", 
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e2e8f0; }
    h1, h2, h3 { color: #0f172a; font-weight: 700; }
    .card-container { background-color: #ffffff; border: 1px solid #e2e8f0; padding: 24px; border-radius: 12px; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); margin-bottom: 20px; }
    .stButton>button { background: #0284c7; color: white; border: none; border-radius: 8px; padding: 0.6rem 1.5rem; font-weight: 600; font-size: 1rem; transition: all 0.2s ease; }
    .stButton>button:hover { background: #0369a1; color: white; }
    .stDownloadButton>button { background: #10b981; color: white; border: none; border-radius: 8px; padding: 0.5rem 1.2rem; font-weight: 600; transition: all 0.2s ease; }
    .stDownloadButton>button:hover { background: #059669; color: white; }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------
# BÚSQUEDA UNIVERSAL Y CACHÉ LOCAL
# ------------------------------------------
gemini_paid_candidates = [st.secrets.get("GEMINI_API_KEY_PAID") if "GEMINI_API_KEY_PAID" in st.secrets else None, os.environ.get("GEMINI_API_KEY_PAID")]
ACTIVE_GEMINI_PAID_KEY = next((k for k in gemini_paid_candidates if k and str(k).strip()), None)

gemini_free_candidates = [st.secrets.get("GEMINI_API_KEY") if "GEMINI_API_KEY" in st.secrets else None, st.secrets.get("GOOGLE_API_KEY") if "GOOGLE_API_KEY" in st.secrets else None, os.environ.get("GEMINI_API_KEY"), os.environ.get("GOOGLE_API_KEY")]
ACTIVE_GEMINI_FREE_KEY = next((k for k in gemini_free_candidates if k and str(k).strip()), None)

BARCODE_CACHE_FILE = "codigos_escaneados_memoria.json"

def load_cache():
    if os.path.exists(BARCODE_CACHE_FILE):
        try:
            with open(BARCODE_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_cache(cache_dict):
    try:
        with open(BARCODE_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache_dict, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

if "barcode_cache" not in st.session_state: st.session_state["barcode_cache"] = load_cache()
if "web_excel_queue" not in st.session_state: st.session_state["web_excel_queue"] = []
if "last_scanned_bc" not in st.session_state: st.session_state["last_scanned_bc"] = ""

def online_barcode_lookup_open(barcode_str):
    """Consulta en caché local o en internet mediante Gemini con alta precisión."""
    cache = st.session_state["barcode_cache"]
    if barcode_str in cache:
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-3.6-flash')
        prompt = (
            f"Identifica con absoluta precisión comercial el producto exacto de bebidas o licores asociado al código de barras EAN/UPC: '{barcode_str}'. "
            "Devuelve un JSON puro con el nombre comercial exacto y presentación oficial en mayúsculas: "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si no lo encuentras, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO EN INTERNET'}."
        )
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        data = json.loads(txt.strip())
        desc = data.get("descripcion", "PRODUCTO DESCONOCIDO EN INTERNET")
        
        # Guardar en caché local
        cache[barcode_str] = desc
        st.session_state["barcode_cache"] = cache
        save_cache(cache)
        
        return desc, "🌐 (Consultado en Internet)"
    except Exception:
        return "PRODUCTO DESCONOCIDO EN INTERNET", "⚠️ (Error de red)"

# ==========================================
# MENÚ Y CONFIGURACIÓN LATERAL
# ==========================================
st.sidebar.markdown("<h3 style='color: #0284c7; text-align: center;'>🌐 Consulta Web</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='text-align: center; color: #64748b; font-size: 0.8rem;'>Módulo Independiente</p>", unsafe_allow_html=True)
st.sidebar.markdown("---")

modulo = st.sidebar.radio("Navegación", ["🌐 Consulta Web de Productos", "📋 Historial y Configuración"])

st.sidebar.markdown("---")
use_gemini_paid_api = st.sidebar.checkbox("💎 Usar Gemini Paid (API de Pago)", value=bool(ACTIVE_GEMINI_PAID_KEY))

# ==========================================
# MÓDULO: CONSULTA WEB DE PRODUCTOS Y EXCEL
# ==========================================
if modulo == "🌐 Consulta Web de Productos":
    try:
        st.markdown("<h2>🌐 Módulo de Consulta Web <span style='color: #0284c7;'>(Caché Local + Lector)</span></h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Escanea con tu lector. Si ya se consultó antes, abrirá al instante (0 seg); si es nuevo, lo buscará en la web.</p>", unsafe_allow_html=True)
        st.markdown("---")

        if "web_bc_input" not in st.session_state: st.session_state["web_bc_input"] = ""
        if "web_desc_result" not in st.session_state: st.session_state["web_desc_result"] = ""
        if "status_msg" not in st.session_state: st.session_state["status_msg"] = ""

        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown("### 📌 Escanear o Ingresar Código de Barra")
        
        def on_barcode_enter():
            val = st.session_state.get("widget_bc_open", "").strip()
            if val and val != st.session_state.get("last_scanned_bc", ""):
                st.session_state["last_scanned_bc"] = val
                clean_bc = str(val).strip()
                with st.spinner(f"Procesando código {clean_bc}..."):
                    found_desc, msg_status = online_barcode_lookup_open(clean_bc)
                st.session_state["web_bc_input"] = clean_bc
                st.session_state["web_desc_result"] = found_desc
                st.session_state["status_msg"] = msg_status

        col_w1, col_w2 = st.columns([2, 1])
        with col_w1:
            input_bc = st.text_input(
                "Código de Barra (EAN / UPC)", 
                placeholder="Escanea aquí con tu lector...", 
                key="widget_bc_open",
                on_change=on_barcode_enter
            )
        with col_w2:
            st.markdown("<br>", unsafe_allow_html=True)
            btn_web_search = st.button("🔍 Consultar Manual")

        if btn_web_search and input_bc.strip():
            clean_bc = str(input_bc).strip()
            with st.spinner(f"Consultando el código {clean_bc}..."):
                found_desc, msg_status = online_barcode_lookup_open(clean_bc)
            st.session_state["web_bc_input"] = clean_bc
            st.session_state["web_desc_result"] = found_desc
            st.session_state["status_msg"] = msg_status

        current_bc = st.session_state.get("web_bc_input", "")
        current_desc = st.session_state.get("web_desc_result", "")
        current_status = st.session_state.get("status_msg", "")

        if current_bc:
            st.markdown("---")
            st.markdown(f"### 📝 Resultado de la Consulta {current_status}")
            
            col_r1, col_r2 = st.columns([2, 1])
            with col_r1:
                final_desc = st.text_input("Descripción Detectada / Editada", value=current_desc, key="input_desc_val")
            with col_r2:
                final_qty = st.number_input("Cantidad", min_value=1, max_value=100000, value=1, step=1, key="input_qty_val")

            if st.button("➕ Agregar a la Vista Previa del Excel"):
                if final_desc.strip():
                    new_entry = {
                        "Código de Barra": str(current_bc),
                        "Descripción": str(final_desc).strip().upper(),
                        "Cantidad": int(final_qty)
                    }
                    st.session_state["web_excel_queue"].append(new_entry)
                    st.success(f"✅ ¡Agregado a la vista previa: **{final_desc}** (Cant: {final_qty})!")
                    st.session_state["web_bc_input"] = ""
                    st.session_state["web_desc_result"] = ""
                    st.session_state["last_scanned_bc"] = ""
                    st.session_state["status_msg"] = ""
                    st.rerun()
                else:
                    st.error("La descripción no puede estar vacía.")

        st.markdown('</div>', unsafe_allow_html=True)

        # VISTA PREVIA EN VIVO
        st.markdown("### 📊 Vista Previa en Vivo del Excel")
        queue = st.session_state.get("web_excel_queue", [])
        
        if queue:
            df_web = pd.DataFrame(queue)
            df_web["Código de Barra"] = df_web["Código de Barra"].astype(str)
            st.dataframe(df_web, use_container_width=True, hide_index=True)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Consulta Web"
            ws.append(["Código de Barra", "Descripción", "Cantidad"])

            for idx_row, row_item in enumerate(queue, start=2):
                ws.append([str(row_item["Código de Barra"]), row_item["Descripción"], int(row_item["Cantidad"])])
                cell = ws.cell(row=idx_row, column=1)
                cell.number_format = '@'

            buffer_excel = io.BytesIO()
            wb.save(buffer_excel)

            col_act1, col_act2 = st.columns(2)
            with col_act1:
                st.download_button(
                    label="📥 Descargar Excel de Consulta Web",
                    data=buffer_excel.getvalue(),
                    file_name=f"Consulta_Web_Productos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_act2:
                if st.button("🗑️ Limpiar Vista Previa"):
                    st.session_state["web_excel_queue"] = []
                    st.rerun()
        else:
            st.info("ℹ️ No hay elementos en la vista previa. Escanea un código arriba para comenzar.")

    except Exception as e:
        st.error("⚠️ Error en el Módulo de Consulta Web:")
        st.exception(e)

elif modulo == "📋 Historial y Configuración":
    st.markdown("<h2>📋 Configuración y Caché Local</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color: #64748b;'>Códigos en memoria caché local: <b>{len(st.session_state.get('barcode_cache', {}))}</b></p>", unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🗑️ Borrar Caché Local de Códigos"):
        st.session_state["barcode_cache"] = {}
        if os.path.exists(BARCODE_CACHE_FILE): os.remove(BARCODE_CACHE_FILE)
        st.success("¡Caché local borrada exitosamente!")
        st.rerun()
