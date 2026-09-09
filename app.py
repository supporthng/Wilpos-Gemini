import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Rápido", page_icon="⚡", layout="wide")

st.title("⚡ Automatizador Rápido de Facturas para WilPOS")
st.markdown("Sube tu factura y la IA procesará todos los productos al instante para entregarte tu Excel con costos sin ITBIS y margen del 25%.")

# Configurar API Key de Gemini desde Streamlit Secrets
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception:
    st.warning("⚠️ No se encontró la GEMINI_API_KEY en los Secrets de Streamlit.")

uploaded_file = st.file_uploader("Sube tu factura (Imagen o PDF)", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    st.success(f"¡Factura cargada: {uploaded_file.name}!")
    
    if st.button("⚡ Procesar Factura Rápido"):
        with st.spinner("Procesando factura a alta velocidad..."):
            try:
                # Usar el modelo rápido optimizado
                model = genai.GenerativeModel('gemini-3.6-flash')
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                
                # Prompt optimizado y ultra conciso para respuesta rápida
                prompt = (
                    "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                    "Calcula el costo unitario sin ITBIS y asegúrate de que sea una respuesta JSON válida sin texto adicional."
                )
                
                response = model.generate_content([
                    {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                    prompt
                ])
                
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]
                raw_text = raw_text.strip()
                
                data_items = json.loads(raw_text)
                
                rows = []
                for item in data_items:
                    costo = float(item.get("costo_sin_itbis", 0))
                    precio_venta = round(costo * 1.25, 2)
                    rows.append({
                        "Código Barra": str(item.get("codigo", "")),
                        "Nombre": str(item.get("descripcion", "")),
                        "Costo Sin ITBIS": costo,
                        "Precio Venta (+25%)": precio_venta,
                        "Stock": int(item.get("stock", 1))
                    })
                
                df_resultado = pd.DataFrame(rows)
                st.success("¡Proceso completado con éxito!")
                st.dataframe(df_resultado, use_container_width=True)
                
                # Generar archivo Excel en formato WilPOS
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Productos"
                ws.append(["Nombre", "Código Barra", "Categoría", "Tipo", "Precio Venta", "Costo", "Stock", "Stock Mínimo", "ITBIS", "Unidad Medida", "Venta Granel", "Cantidad Empaque", "Precio Variable", "Descuento %", "Descuento Monto", "Precio Especial", "Descuento Activo", "Descuento Nota"])
                
                for item_dict in data_items:
                    costo = float(item_dict.get("costo_sin_itbis", 0))
                    pv = round(costo * 1.25, 2)
                    ws.append([
                        str(item_dict.get("descripcion", "")),
                        str(item_dict.get("codigo", "")),
                        "General",
                        "producto",
                        pv,
                        costo,
                        int(item_dict.get("stock", 1)),
                        5,
                        0.18,
                        "unidad",
                        "No",
                        int(item_dict.get("empaque", 1)),
                        "No",
                        0,
                        0,
                        None,
                        "No",
                        None
                    ])
                
                output = io.BytesIO()
                wb.save(output)
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Descargar Excel WilPOS Actualizado",
                    data=excel_data,
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Error en el procesamiento rápido: {e}")
