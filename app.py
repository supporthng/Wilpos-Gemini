import streamlit as st
import openpyxl
import pandas as pd
io_module = __import__('io')
import json
import os
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador con Plantilla", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Facturas para WilPOS (Plantilla Oficial)")
st.markdown("Sube tu factura. La app extraerá los productos mostrando **solo el nombre y el tamaño**, calculará la fórmula (**Costo + 25% + 18% ITBIS** con redondeo a **múltiplos de 5**), mantendrá los **ceros a la izquierda** en los códigos de barras y usará la **Plantilla Oficial de WilPOS**.")

# Configurar claves desde Streamlit Secrets (Soporta múltiples cuentas gratuitas y pago opcional)
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

if uploaded_file is not None:
    st.success(f"¡Factura cargada: {uploaded_file.name}!")
    
    if st.session_state["quota_exceeded"]:
        st.warning("⚠️ **Se han agotado todas las solicitudes de las cuentas gratuitas principales.**")
        confirm_paid = st.checkbox("¿Deseas procesar esta factura utilizando la versión de pago (o tienes otra clave adicional)?")
        
        if confirm_paid:
            if st.button("🚀 Continuar con Versión Alternativa / Pago"):
                with st.spinner("Procesando factura..."):
                    try:
                        genai.configure(api_key=paid_api_key)
                        model_paid = genai.GenerativeModel('gemini-3.6-flash')
                        
                        uploaded_file.seek(0)
                        file_bytes = uploaded_file.read()
                        
                        prompt = (
                            "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                            "REGLA PARA DESCRIPCIÓN: Limpia la descripción del producto para que solo incluya el nombre principal y su tamaño/gramaje (ejemplo: 'RUFFLES CHEDDAR 120G', eliminando términos como TA, CS, 16X1). "
                            "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' estrictamente como texto (string). NO elimines los ceros a la izquierda (por ejemplo, '0300055292'). "
                            "Nota importante: 'stock' y 'empaque' deben ser valores numéricos enteros. Si no hay stock, pon 1. "
                            "Calcula el costo unitario sin ITBIS y asegúrate de que sea una respuesta JSON válida sin texto adicional."
                        )
                        
                        response = model_paid.generate_content([
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
                        st.session_state["quota_exceeded"] = False
                        st.success("✅ ¡Factura procesada exitosamente!")
                        
                        rows_preview = []
                        for idx, item in enumerate(data_items, start=1):
                            costo = safe_float(item.get("costo_sin_itbis", 0))
                            raw_pv = (costo * 1.25) * 1.18
                            precio_venta = round_to_nearest_5(raw_pv)
                            stock_val = safe_int(item.get("stock", 1), 1)
                            codigo_barras = str(item.get("codigo", "")).strip()
                            
                            rows_preview.append({
                                "No.": idx,
                                "Código Barra": codigo_barras,
                                "Nombre": str(item.get("descripcion", "")),
                                "Costo Sin ITBIS": costo,
                                "Precio Venta (Múltiplos de 5)": precio_venta,
                                "Stock": stock_val
                            })
                        
                        df_resultado = pd.DataFrame(rows_preview)
                        st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                        
                        template_path = "Plantilla_Inventario_WilPOS_2.xlsx"
                        if not os.path.exists(template_path):
                            template_path = "Plantilla_Inventario_WilPOS.xlsx"
                            
                        if os.path.exists(template_path):
                            wb = openpyxl.load_workbook(template_path)
                            ws_prod = wb['Productos']
                            ws_prod.delete_rows(2, ws_prod.max_row)
                        else:
                            wb = openpyxl.Workbook()
                            ws_prod = wb.active
                            ws_prod.title = "Productos"
                            ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                        
                        for item_dict in data_items:
                            costo = safe_float(item_dict.get("costo_sin_itbis", 0))
                            raw_pv = (costo * 1.25) * 1.18
                            pv = round_to_nearest_5(raw_pv)
                            empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                            stock_val = safe_int(item_dict.get("stock", 1), 1)
                            codigo_barras = str(item_dict.get("codigo", "")).strip()
                            
                            ws_prod.append([
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
                            ])
                            ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                        
                        output = io_module.BytesIO()
                        wb.save(output)
                        excel_data = output.getvalue()
                        
                        st.download_button(
                            label="📥 Descargar Excel Plantilla WilPOS Actualizada",
                            data=excel_data,
                            file_name="Inventario_WilPOS_Actualizado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as err_paid:
                        st.error(f"Error al procesar: {err_paid}")
    else:
        if st.button("🚀 Procesar Factura con Plantilla Oficial"):
            with st.spinner("Analizando factura..."):
                data_items = None
                success_msg = ""
                
                # Intento 1: Cuenta Gratuita Principal (GEMINI_API_KEY)
                try:
                    if not free_key_1:
                        raise Exception("No free key 1")
                    
                    genai.configure(api_key=free_key_1)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    
                    uploaded_file.seek(0)
                    file_bytes = uploaded_file.read()
                    
                    prompt = (
                        "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                        "REGLA PARA DESCRIPCIÓN: Limpia la descripción del producto para que solo incluya el nombre principal y su tamaño/gramaje (ejemplo: 'RUFFLES CHEDDAR 120G', eliminando términos como TA, CS, 16X1). "
                        "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' estrictamente como texto (string). NO elimines los ceros a la izquierda (por ejemplo, '0300055292'). "
                        "Nota importante: 'stock' y 'empaque' deben ser valores numéricos enteros. Si no hay stock, pon 1. "
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
                    success_msg = "✅ ¡Factura procesada exitosamente usando la **Cuenta Gratuita #1**!"
                    
                except Exception as e1:
                    err_msg1 = str(e1)
                    # Si falla la cuota 1, intentamos automáticamente con la Cuenta Gratuita #2 (GEMINI_API_KEY_2)
                    if ("429" in err_msg1 or "Quota exceeded" in err_msg1 or "No free key 1" in err_msg1) and free_key_2:
                        try:
                            genai.configure(api_key=free_key_2)
                            model2 = genai.GenerativeModel('gemini-3.6-flash')
                            
                            uploaded_file.seek(0)
                            file_bytes = uploaded_file.read()
                            
                            prompt = (
                                "Extrae todos los productos de esta factura en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                                "REGLA PARA DESCRIPCIÓN: Limpia la descripción del producto para que solo incluya el nombre principal y su tamaño/gramaje (ejemplo: 'RUFFLES CHEDDAR 120G', eliminando términos como TA, CS, 16X1). "
                                "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' estrictamente como texto (string). NO elimines los ceros a la izquierda (por ejemplo, '0300055292'). "
                                "Nota importante: 'stock' y 'empaque' deben ser valores numéricos enteros. Si no hay stock, pon 1. "
                                "Calcula el costo unitario sin ITBIS y asegúrate de que sea una respuesta JSON válida sin texto adicional."
                            )
                            
                            response = model2.generate_content([
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
                            success_msg = "✅ ¡Factura procesada exitosamente rotando a la **Cuenta Gratuita #2**!"
                        except Exception as e2:
                            err_msg2 = str(e2)
                            if "429" in err_msg2 or "Quota exceeded" in err_msg2:
                                st.session_state["quota_exceeded"] = True
                                st.rerun()
                            else:
                                st.error(f"Error con la Cuenta Gratuita #2: {e2}")
                    elif "429" in err_msg1 or "Quota exceeded" in err_msg1:
                        st.session_state["quota_exceeded"] = True
                        st.rerun()
                    else:
                        st.error(f"Ocurrió un error al procesar con la IA: {e1}")
                
                if data_items:
                    st.success(success_msg)
                    rows_preview = []
                    for idx, item in enumerate(data_items, start=1):
                        costo = safe_float(item.get("costo_sin_itbis", 0))
                        raw_pv = (costo * 1.25) * 1.18
                        precio_venta = round_to_nearest_5(raw_pv)
                        stock_val = safe_int(item.get("stock", 1), 1)
                        codigo_barras = str(item.get("codigo", "")).strip()
                        
                        rows_preview.append({
                            "No.": idx,
                            "Código Barra": codigo_barras,
                            "Nombre": str(item.get("descripcion", "")),
                            "Costo Sin ITBIS": costo,
                            "Precio Venta (Múltiplos de 5)": precio_venta,
                            "Stock": stock_val
                        })
                    
                    df_resultado = pd.DataFrame(rows_preview)
                    st.dataframe(df_resultado, use_container_width=True, hide_index=True)
                    
                    template_path = "Plantilla_Inventario_WilPOS_2.xlsx"
                    if not os.path.exists(template_path):
                        template_path = "Plantilla_Inventario_WilPOS.xlsx"
                        
                    if os.path.exists(template_path):
                        wb = openpyxl.load_workbook(template_path)
                        ws_prod = wb['Productos']
                        ws_prod.delete_rows(2, ws_prod.max_row)
                    else:
                        wb = openpyxl.Workbook()
                        ws_prod = wb.active
                        ws_prod.title = "Productos"
                        ws_prod.append(['Nombre', 'Código Barra', 'Categoría', 'Tipo', 'Precio Venta', 'Costo', 'Stock', 'Stock Mínimo', 'ITBIS', 'Unidad Medida', 'Venta Granel', 'Cantidad Empaque', 'Precio Variable', 'Descuento %', 'Descuento Monto', 'Precio Especial', 'Descuento Activo', 'Descuento Nota'])
                    
                    for item_dict in data_items:
                        costo = safe_float(item_dict.get("costo_sin_itbis", 0))
                        raw_pv = (costo * 1.25) * 1.18
                        pv = round_to_nearest_5(raw_pv)
                        empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                        stock_val = safe_int(item_dict.get("stock", 1), 1)
                        codigo_barras = str(item_dict.get("codigo", "")).strip()
                        
                        ws_prod.append([
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
                        ])
                        ws_prod.cell(row=ws_prod.max_row, column=2).number_format = '@'
                    
                    output = io_module.BytesIO()
                    wb.save(output)
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 Descargar Excel Plantilla WilPOS Actualizada",
                        data=excel_data,
                        file_name="Inventario_WilPOS_Actualizado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
