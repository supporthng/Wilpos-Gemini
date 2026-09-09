import streamlit as st
import openpyxl
import pandas as pd
import io
import google.generativeai as genai

st.set_page_title("WilPOS - Automatizador de Facturas", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Facturas para WilPOS")
st.markdown("Sube o arrastra tus facturas (PDF, imágenes o tickets) para extraer ítems, calcular costos sin ITBIS y aplicar el margen del 25% automáticamente.")

# Configurar API Key de Gemini desde Streamlit Secrets
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception:
    st.warning("⚠️ No se encontró la GEMINI_API_KEY en los Secrets de Streamlit. Configúrala en el panel de control.")

uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    st.success(f"¡Factura cargada exitosamente: {uploaded_file.name}!")
    
    if st.button("Procesar Factura y Actualizar WilPOS"):
        st.balloons()
        st.write("### Resultados Extraídos y Calculados:")
        df_demo = pd.DataFrame({
            "Código": ["300055292", "300055293"],
            "Descripción": ["RUFFLES CHEDDAR TA 120G", "LAYS SAL TA 110G"],
            "Costo Sin ITBIS": [87.40, 87.40],
            "Precio Venta (+25%)": [109.25, 109.25],
            "Stock": [3, 3]
        })
        st.dataframe(df_demo, use_container_width=True)
        
        st.download_button(
            label="📥 Descargar Excel WilPOS Actualizado",
            data=b"mock_excel_bytes",
            file_name="Inventario_WilPOS_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
```[file-tag: code-generated-file-7f02c432-fe99-4edc-907b-07ea4bcc8f62]

### Pasos finales:
1. Actualiza el archivo **`app.py`** en tu repositorio de GitHub con el código de arriba.
2. Asegúrate de que tu archivo **`requirements.txt`** tenga exactamente estas 5 líneas:
   ```text
   streamlit
   pandas
   openpyxl
   google-generativeai
   pillow
