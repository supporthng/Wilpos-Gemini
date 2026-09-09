import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Rápido", page_icon="⚡", layout="wide")

st.title("⚡ Automatizador Rápido de Facturas para WilPOS")
st.markdown("Sube tu factura y la IA procesará los productos aplicando la fórmula (**Costo + 25% + 18% ITBIS**) con redondeo automático a **múltiplos de 5**.")

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
    
    if st.button("⚡ Procesar con Precios Redondeados (Múltiplos de 5)"):
        with st.spinner("Procesando factura y aplicando redondeo a múltiplos de 5..."):
            try:
                model = genai.GenerativeModel('gemini-3.6-flash')
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                
                prompt = (
                    "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                    "Nota importante: 'stock' y 'empaque' deben ser valores numéricos enteros (ejemplo: 3, 12, 1). Si no hay stock especificado, pon 1. "
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
                    costo = safe_float(item.get("costo_sin_itbis", 0))
                    # Fórmula: (Costo * 1.25) * 1.18, redondeado al múltiplo de 5 más cercano
                    raw_pv = (costo * 1.25) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    stock_val = safe_int(item.get("stock", 1), 1)
                    rows.append({
                        "Código Barra": str(item.get("codigo", "")),
                        "Nombre": str(item.get("descripcion", "")),
                        "Costo Sin ITBIS": costo,
                        "Precio Venta (Redondeado a Múltiplos de 5)": precio_venta,
                        "Stock": stock_val
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
                    costo = safe_float(item_dict.get("costo_sin_itbis", 0))
                    raw_pv = (costo * 1.25) * 1.18
                    pv = round_to_nearest_5(raw_pv)
                    empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                    stock_val = safe_int(item_dict.get("stock", 1), 1)
                    
                    ws.append([
                        str(item_dict.get("descripcion", "")),
                        str(item_dict.get("codigo", "")),
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
                    ])
                
                output = io.BytesIO()
                wb.save(output)
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Descargar Excel WilPOS (Precios en Múltiplos de 5)",
                    data=excel_data,
                    file_name="Inventario_WilPOS_Actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Error en el procesamiento: {e}")
```[file-tag: code-generated-file-16e12e11-6930-4055-a663-6db7b9ed55a7]

Haz un *Reboot* en Streamlit Cloud tras guardar este cambio y tus precios de venta se generarán directamente redondeados en múltiplos de 5.
