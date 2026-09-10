import streamlit as st
import openpyxl
import pandas as pd
io_module = __import__('io')
import json
import os
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Inteligente de Facturas", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Facturas para WilPOS (Detección Inteligente de Empaques)")
st.markdown("Sube cualquier factura. La IA detectará automáticamente el tipo de empaque (Caja, Paquete, Lata, Botella), calculará las cantidades y el **costo unitario exacto sin ITBIS**, aplicará la fórmula del **25% de margen + 18% ITBIS** con redondeo a **múltiplos de 5**, y generará el Excel usando la **Plantilla Oficial de WilPOS**.")

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
        st.warning("⚠️ **Se han agotado las solicitudes gratuitas de ambas cuentas de Gemini.**")
        confirm_paid = st.checkbox("¿Deseas procesar esta factura utilizando la versión de pago?")
        
        if confirm_paid:
            if st.button("🚀 Continuar con la Versión de Pago"):
                with st.spinner("Procesando factura con la versión de pago..."):
                    try:
                        genai.configure(api_key=paid_api_key)
                        model_paid = genai.GenerativeModel('gemini-3.6-flash')
                        
                        uploaded_file.seek(0)
                        file_bytes = uploaded_file.read()
                        
                        prompt = (
                            "Analiza esta factura detectando los diferentes tipos de empaques (Caja, Paquete, Botella, Lata, Unidad, etc.), sus cantidades y el factor de conversión. "
                            "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                            "REGLA DE ORO PARA EL COSTO UNITARIO: Identifica el valor total de la línea y el ITBIS. Calcula el valor neto sin ITBIS (Valor Total - ITBIS). Luego, detecta cuántas unidades individuales componen el empaque (ej. Caja-12 = 12 unidades, Paquete-24 = 24 unidades) y divide el neto entre el total de unidades para obtener el 'costo_sin_itbis' por unidad individual exacta. "
                            "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML', eliminando términos de empaque masivo como CS, TA, 16X1). "
                            "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' como texto (string) preservando todos los ceros a la izquierda. "
                            "Stock y empaque deben ser enteros numéricos. Respuesta JSON válida sin texto adicional."
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
                        st.success("✅ ¡Factura procesada exitosamente con la versión de pago!")
                        
                        rows_preview = []
                        for idx, item in enumerate(data_items, start=1):
                            costo = safe_float(item.get("costo_sin_itbis", 0))
                            raw_pv = (costo * 1.25) * 1.18
                            precio_venta = round_to_nearest_5(raw_pv)
                            stock_val = safe_int(item.get("stock", 1), 1)
                            empaque_val = safe_int(item.get("empaque", 1), 1)
                            codigo_barras = str(item.get("codigo", "")).strip()
                            
                            rows_preview.append({
                                "No.": idx,
                                "Código Barra": codigo_barras,
                                "Nombre": str(item.get("descripcion", "")),
                                "Empaque": empaque_val,
                                "Costo Unit. Sin ITBIS": costo,
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
                        st.error(f"Error al procesar con la versión de pago: {err_paid}")
    else:
        if st.button("🚀 Procesar Factura con Plantilla Oficial"):
            with st.spinner("Analizando factura y calculando costos unitarios..."):
                data_items = None
                success_msg = ""
                
                prompt_text = (
                    "Analiza esta factura detectando los diferentes tipos de empaques (Caja, Paquete, Botella, Lata, Unidad, etc.), sus cantidades y el factor de conversión. "
                    "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'costo_sin_itbis', 'empaque', 'stock'). "
                    "REGLA DE ORO PARA EL COSTO UNITARIO: Identifica el valor total de la línea y el ITBIS. Calcula el valor neto sin ITBIS (Valor Total - ITBIS). Luego, detecta cuántas unidades individuales componen el empaque (ej. Caja-12 = 12 unidades, Paquete-24 = 24 unidades) y divide el neto entre el total de unidades para obtener el 'costo_sin_itbis' por unidad individual exacta. "
                    "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML', eliminando términos de empaque masivo como CS, TA, 16X1). "
                    "REGLA CRÍTICA PARA CÓDIGOS: Trata el campo 'codigo' como texto (string) preservando todos los ceros a la izquierda. "
                    "Stock y empaque deben ser enteros numéricos. Respuesta JSON válida sin texto adicional."
                )
                
                # Intento con Cuenta Gratuita #1
                try:
                    if not free_key_1:
                        raise Exception("No free key 1")
                    
                    genai.configure(api_key=free_key_1)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    
                    uploaded_file.seek(0)
                    file_bytes = uploaded_file.read()
                    
                    response = model.generate_content([
                        {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                        prompt_text
                    ])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    data_items = json.loads(raw_text)
                    success_msg = "✅ ¡Factura procesada con éxito usando la **Cuenta Gratuita #1**!"
                    
                except Exception as e1:
                    err_msg1 = str(e1)
                    # Rotación automática a Cuenta Gratuita #2
                    if ("429" in err_msg1 or "Quota exceeded" in err_msg1 or "No free key 1" in err_msg1) and free_key_2:
                        try:
                            genai.configure(api_key=free_key_2)
                            model2 = genai.GenerativeModel('gemini-3.6-flash')
                            
                            uploaded_file.seek(0)
                            file_bytes = uploaded_file.read()
                            
                            response = model2.generate_content([
                                {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                                prompt_text
                            ])
                            
                            raw_text = response.text.strip()
                            if raw_text.startswith("```json"):
                                raw_text = raw_text[7:]
                            if raw_text.endswith("```"):
                                raw_text = raw_text[:-3]
                            raw_text = raw_text.strip()
                            
                            data_items = json.loads(raw_text)
                            success_msg = "✅ ¡Factura procesada con éxito rotando a la **Cuenta Gratuita #2**!"
                        except Exception as e2:
                            err_msg2 = str(e2)
                            if "429" in err_msg2 or "Quota exceeded" in err_msg2:
                                st.session_state["quota_exceeded"] = True
                                st.rerun()
                            else:
                                st.error(f"Error con Cuenta Gratuita #2: {e2}")
                    elif "429" in err_msg1 or "Quota exceeded" in err_msg1:
                        st.session_state["quota_exceeded"] = True
                        st.rerun()
                    else:
                        st.error(f"Ocurrió un error con la IA: {e1}")
                
                if data_items:
                    st.success(success_msg)
                    rows_preview = []
                    for idx, item in enumerate(data_items, start=1):
                        costo = safe_float(item.get("costo_sin_itbis", 0))
                        raw_pv = (costo * 1.25) * 1.18
                        precio_venta = round_to_nearest_5(raw_pv)
                        stock_val = safe_int(item.get("stock", 1), 1)
                        empaque_val = safe_int(item.get("empaque", 1), 1)
                        codigo_barras = str(item.get("codigo", "")).strip()
                        
                        rows_preview.append({
                            "No.": idx,
                            "Código Barra": codigo_barras,
                            "Nombre": str(item.get("descripcion", "")),
                            "Empaque": empaque_val,
                            "Costo Unit. Sin ITBIS": costo,
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
