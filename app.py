import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador de Facturas", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Facturas para WilPOS")
st.markdown("Sube o arrastra tus facturas (PDF, imágenes o tickets) para extraer **todos los ítems**, calcular costos sin ITBIS y aplicar el margen del 25% automáticamente.")

# Configurar API Key de Gemini desde Streamlit Secrets
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception:
    st.warning("⚠️ No se encontró la GEMINI_API_KEY en los Secrets de Streamlit. Configúrala en el panel de control.")

uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    st.success(f"¡Factura cargada exitosamente: {uploaded_file.name}!")
    
    try:
        image = Image.open(uploaded_file)
        st.image(image, caption="Factura Cargada", use_column_width=True)
    except Exception:
        pass

    if st.button("🚀 Procesar Factura con IA y Actualizar WilPOS"):
        with st.spinner("Analizando factura, extrayendo productos y calculando precios con IA (Gemini 3.6 Flash)..."):
            try:
                # Usar el modelo recomendado: gemini-3.6-flash
                model = genai.GenerativeModel('gemini-3.6-flash')
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                
                prompt = (
                    "Analiza esta factura detalladamente. Extrae TODOS los productos de la tabla. "
                    "Para cada producto, devuelve estrictamente un arreglo JSON válido (sin formato de bloque de código markdown adicional si es posible, o puro JSON) "
                    "que contenga una lista de objetos con las siguientes claves exactas: "
                    "'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'. "
                    "Calcula el costo unitario sin ITBIS (si incluye 18% de ITBIS desglósalo, o si es neto úsalo directo por unidad/caja según corresponda)."
                )
                
                response = model.generate_content([
                    {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                    prompt
                ])
                
                # Limpiar texto de respuesta para extraer JSON
                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]
                raw_text = raw_text.strip()
                
                data_items = json.loads(raw_text)
                
                # Construir DataFrame
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
                st.success("¡Factura procesada exitosamente con IA!")
                st.dataframe(df_resultado, use_container_width=True)
                
                # Generar Excel real con openpyxl en formato WilPOS
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
                        "Snacks / Licores",
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
                    label="📥 Descargar Excel WilPOS Completo Actualizado",
                    data=excel_data,
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Ocurrió un error al procesar con la IA: {e}")
