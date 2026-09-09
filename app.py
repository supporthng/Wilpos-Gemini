import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Rápido", page_icon="⚡", layout="wide")

st.title("⚡ Automatizador Rápido de Facturas para WilPOS")
st.markdown("Sube tu factura y la IA procesará los productos aplicando la fórmula (**Costo + 25% + 18% ITBIS**) con redondeo automático a **múltiplos de 5** y conservación de **ceros a la izquierda** en los códigos.")

# Configurar API Key de Gemini desde Streamlit Secrets
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
except Exception:
    st.warning("⚠️ No se encontró la GEMINI_API_KEY en los Secrets de Streamlit.")

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

if uploaded_file is not None:
    st.success(f"¡Factura cargada: {uploaded_file.name}!")
    
    if st.button("⚡ Procesar Factura (Conservando Ceros y Redondeando)"):
        with st.spinner("Procesando factura y asegurando precios y códigos exactos..."):
            try:
                model = genai.GenerativeModel('gemini-3.6-flash')
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                
                prompt = (
                    "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                    "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' estrictamente como texto (string). NO elimines los ceros a la izquierda (por ejemplo, si el código es '05455458444', debe mantenerse completo con su cero inicial). "
                    "Nota importante: 'stock' y 'empaque' deben ser valores numéricos enteros. Si no hay stock especificado, pon 1. "
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
                for idx, item in enumerate(data_items, start=1):
                    costo = safe_float(item.get("costo_sin_itbis", 0))
                    # Fórmula: (Costo * 1.25) * 1.18, redondeado al múltiplo de 5 más cercano
                    raw_pv = (costo * 1.25) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = safe_int(item.get("stock", 1), 1)
                    
                    codigo_barras = str(item.get("codigo", "")).strip()
                    
                    rows.append({
                        "No.": idx,
                        "Código Barra": codigo_barras,
                        "Nombre": str(item.get("descripcion", "")),
                        "Costo Sin ITBIS": costo,
                        "Precio Venta (Múltiplos de 5)": precio_venta,
                        "Stock": stock_val
                    })
                
                df_resultado = pd.DataFrame(rows)
                st.success("¡Proceso completado con éxito!")
                st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                
                # Generar archivo Excel en formato WilPOS
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Productos"
                ws.append(["No.", "Nombre", "Código Barra", "Categoría", "Tipo", "Precio Venta", "Costo", "Stock", "Stock Mínimo", "ITBIS", "Unidad Medida", "Venta Granel", "Cantidad Empaque", "Precio Variable", "Descuento %", "Descuento Monto", "Precio Especial", "Descuento Activo", "Descuento Nota"])
                
                for idx, item_dict in enumerate(data_items, start=1):
                    costo = safe_float(item_dict.get("costo_sin_itbis", 0))
                    raw_pv = (costo * 1.25) * 1.18
                    pv = round_to_nearest_5(raw_pv)
                    empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                    stock_val = safe_int(item_dict.get("stock", 1), 1)
                    codigo_barras = str(item_dict.get("codigo", "")).strip()
                    
                    row_cells = [
                        idx,
                        str(item_dict.get("descripcion", "")),
                        codigo_barras,
                        "General",
                        "producto",
                        pv,
                        costo,
                        stock_val,
                        5,
                        0.18,
                        "unidad",
                        "No",
                        empaque_val,
                        "No",
                        0,
                        0,
                        None,
                        "No",
                        None
                    ]
                    ws.append(row_cells)
                    # Forzar formato texto en la celda del código de barras en Excel
                    ws.cell(row=ws.max_row, column=3).number_format = '@'
                
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
                st.error(f"Error en el procesamiento: {e}")
