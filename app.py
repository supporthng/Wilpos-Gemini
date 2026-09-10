import streamlit as st
import openpyxl
import pandas as pd
io_module = __import__('io')
import json
import os
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Inteligente de Facturas", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Facturas para WilPOS (Cálculo Exacto de Stock y Costos)")
st.markdown("Sube tu factura. La IA detectará los empaques, calculará el costo unitario sin ITBIS y el **stock total correcto multiplicando la cantidad comprada por el tamaño del empaque**, aplicando la fórmula de WilPOS (**25% margen + 18% ITBIS** con redondeo a **múltiplos de 5**).")

free_key_1 = st.secrets.get("GEMINI_API_KEY", "")
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", "")
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", "")

uploaded_file = st.file_uploader("Sube tu factura (Imagen o PDF)", type=["pdf", "png", "jpg", "jpeg"])

def safe_int(val, default=1):
    try:
        return int(float(val))
    except Exception:
        return default

def safe_float(val, default=0.0):
    try:
        return float(val)
    except Exception:
        return default

def round_to_nearest_5(x):
    return round(round(x / 5) * 5, 2)

if "quota_exceeded" not in st.session_state:
    st.session_state["quota_exceeded"] = False

# Definir la ventana emergente (modal centrado) usando st.dialog
@st.dialog("⚠️ Límite de Cuota Gratuita Alcanzado")
def paid_confirmation_dialog():
    st.write("Se han agotado las solicitudes gratuitas de **ambas cuentas** de Gemini.")
    st.write("¿Deseas procesar esta factura utilizando la **versión de pago**?")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 Sí, usar Versión de Pago", type="primary"):
            st.session_state["use_paid_now"] = True
            st.rerun()
    with col2:
        if st.button("❌ Cancelar"):
            st.session_state["quota_exceeded"] = False
            st.rerun()

if "use_paid_now" not in st.session_state:
    st.session_state["use_paid_now"] = False

if uploaded_file is not None:
    st.success(f"¡Factura cargada: {uploaded_file.name}!")
    
    if st.session_state["quota_exceeded"]:
        paid_confirmation_dialog()
        
    if st.session_state["use_paid_now"]:
        with st.spinner("Procesando factura con la versión de pago..."):
            try:
                genai.configure(api_key=paid_api_key)
                model_paid = genai.GenerativeModel('gemini-3.6-flash')
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                
                prompt = (
                    "Analiza esta factura detalladamente. Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (la cantidad comprada de cajas/paquetes/unidades, ej: 2, 4, 5, 6), 'empaque' (cuántas unidades individuales trae cada caja o paquete, ej: 12, 24, 10, o 1 si es suelto), y el valor total e ITBIS. "
                    "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'cantidad', 'empaque', 'costo_sin_itbis'). "
                    "REGLA DE ORO PARA EL COSTO UNITARIO: Calcula el valor neto sin ITBIS (Valor Total - ITBIS) y divídelo entre el total de unidades individuales (cantidad * empaque) para obtener el 'costo_sin_itbis' por unidad exacta. "
                    "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML'). "
                    "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' como texto (string) preservando ceros a la izquierda. "
                    "Respuesta JSON válida sin texto adicional."
                )
                
                response = model_paid.generate_content([
                    {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                    prompt
                ])
                
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.endswith("
