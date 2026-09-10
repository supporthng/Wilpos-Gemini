import io
import json
import os
import hashlib
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

# Funciones auxiliares de cálculo
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

# Inicializar control de documentos procesados para evitar duplicados
if "processed_signatures" not in st.session_state:
    st.session_state["processed_signatures"] = set()

if "quota_exceeded" not in st.session_state:
    st.session_state["quota_exceeded"] = False

if "use_paid_now" not in st.session_state:
    st.session_state["use_paid_now"] = False

def paid_confirmation_dialog():
    st.warning("⚠️ Se ha agotado la cuota de las cuentas gratuitas de Gemini (Error 429 / Quota Exceeded).")
    if st.button("💳 Continuar usando la Versión de Pago (API Key de Pago)"):
        st.session_state["use_paid_now"] = True
        st.session_state["quota_exceeded"] = False
        st.rerun()

# ==========================================
# MENÚ DE NAVEGACIÓN LATERAL
# ==========================================
st.sidebar.title("Menú de Navegación")
modulo = st.sidebar.radio(
    "Selecciona el Módulo",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

# Botón para limpiar historial de duplicados si el usuario lo desea
if st.sidebar.button("🧹 Restablecer Memoria de Duplicados"):
    st.session_state["processed_signatures"] = set()
    st.sidebar.success("¡Memoria de duplicados reiniciada!")

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura. El sistema cuenta con **control anti-duplicados** y cálculo exacto de costos unitarios netos sin ITBIS.")

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")

        if st.session_state["quota_exceeded"]:
            paid_confirmation_dialog()

        if st.button("🚀 Procesar Factura con Control Anti-Duplicados") or st.session_state["use_paid_now"]:
            with st.spinner("Analizando factura, calculando costos reales y verificando duplicidad..."):
                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (la cantidad comprada, ej: 10, 20), 'empaque' (unidades por empaque, ej: 1 si es por unidad directa), y 'costo_sin_itbis'. "
                    "REGLA CRÍTICA PARA EL COSTO: Los precios mostrados en las líneas de la factura suelen incluir impuestos o representar el monto total de la línea. Debes calcular rigurosamente el costo unitario real SIN ITBIS por cada unidad individual (descontando el ITBIS global si aplica y dividiendo entre cantidad * empaque). "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA PARA CÓDIGOS DE BARRAS: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )
                
                file_bytes = uploaded_file.read()
                file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg'
                
                parsed_data = None
                success_msg = ""
                
                active_key = paid_api_key if st.session_state["use_paid_now"] else free_key_1
                
                try:
                    if not active_key and not st.session_state["use_paid_now"]:
                        active_key = free_key_2
                        
                    genai.configure(api_key=active_key if active_key else paid_api_key)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    
                    uploaded_file.seek(0)
                    file_bytes = uploaded_file.read()
                    
                    response = model.generate_content([
                        {'mime_type': file_type, 'data': file_bytes},
                        prompt_text
                    ])
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()
                    
                    parsed_data = json.loads(raw_text)
                    st.session_state["use_paid_now"] = False
                    success_msg = "✅ ¡Factura procesada con éxito!"
                    
                except Exception as e:
                    err_str = str(e)
                    if ("429" in err_str or "Quota exceeded" in err_str) and not st.session_state["use_paid_now"]:
                        try:
                            genai.configure(api_key=free_key_2 if free_key_2 else paid_api_key)
                            model2 = genai.GenerativeModel('gemini-3.6-flash')
                            uploaded_file.seek(0)
                            response = model2.generate_content([
                                {'mime_type': file_type, 'data': uploaded_file.read()},
                                prompt_text
                            ])
                            raw_text = response.text.strip()
                            if raw_text.startswith("```json"):
                                raw_text = raw_text[7:]
                            if raw_text.endswith("```"):
                                raw_text = raw_text[:-3]
                            parsed_data = json.loads(raw_text.strip())
                            success_msg = "✅ ¡Factura procesada usando el respaldo gratuito #2!"
                        except Exception:
                            st.session_state["quota_exceeded"] = True
                            st.rerun()
                    else:
                        st.error(f"Error al procesar: {e}")

                if parsed_data:
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()

                    if doc_signature in st.session_state["processed_signatures"]:
                        st.error("🚨 **¡ADVERTENCIA DE FACTURA DUPLICADA!**")
                        data_msg = (
                            f"Este documento ya fue procesado anteriormente.\n\n"
                            f"- **Emisor RNC:** `{rnc_emisor}`\n"
                            f"- **No. Documento:** `{num_doc}`\n"
                            f"- **Fecha:** `{fecha_doc}`\n"
                            f"- **Total:** `{total_doc}`\n\n"
                            f"El sistema ha bloqueado la carga para evitar duplicidad en el inventario de WilPOS."
                        )
                        st.warning(data_msg)
                    else:
                        st.session_state["processed_signatures"].add(doc_signature)
                        st.success(success_msg)
                        
                        data_items = parsed_data.get("items", [])
                        rows_preview = []
                        for idx, item in enumerate(data_items, start=1):
                            costo = safe_float(item.get("costo_sin_itbis", 0))
                            raw_pv = (costo * 1.25) * 1.18
                            precio_venta = round_to_nearest_5(raw_pv)
                            cant_comprada = safe_int(item.get("cantidad", 1), 1)
                            empaque_val = safe_int(item.get("empaque", 1), 1)
                            stock_val = cant_comprada * empaque_val
                            codigo_barras = str(item.get("codigo", "")).strip()
                            
                            rows_preview.append({
                                "No.": idx,
                                "Código Barra": codigo_barras,
                                "Nombre": str(item.get("descripcion", "")),
                                "Cant. Compra": cant_comprada,
                                "Empaque": empaque_val,
                                "Stock Total": stock_val,
                                "Costo Unit. Sin ITBIS": costo,
                                "Precio Venta (M5)": precio_venta
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
                            cant_comprada = safe_int(item_dict.get("cantidad", 1), 1)
                            empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                            stock_val = cant_comprada * empaque_val
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
                        
                        output = io.BytesIO()
                        wb.save(output)
                        excel_data = output.getvalue()
                        
                        st.download_button(
                            label="📥 Descargar Excel Plantilla WilPOS Actualizada",
                            data=excel_data,
                            file_name="Inventario_WilPOS_Actualizado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE)
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.title("📂 Procesador por Lotes con Control Anti-Duplicados")
    st.markdown("Sube varias facturas o cotizaciones. El sistema filtrará automáticamente cualquier documento duplicado evaluando su firma única.")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        if st.button("🚀 Procesar Lote Evaluando Duplicados", type="primary"):
            all_consolidated_items = []
            duplicate_count = 0
            processed_in_this_batch = set()

            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}...")

                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada), 'empaque' (unidades por empaque), y 'costo_sin_itbis' (calculado por unidad real sin ITBIS). "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )

                file_bytes = file.read()
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                
                parsed_data = None
                
                try:
                    if free_key_1:
                        genai.configure(api_key=free_key_1)
                        model_b1 = genai.GenerativeModel('gemini-3.6-flash')
                        response = model_b1.generate_content([
                            {'mime_type': file_type, 'data': file_bytes},
                            prompt_text
                        ])
                        raw_txt = response.text.strip()
                        if raw_txt.startswith("```json"):
                            raw_txt = raw_txt[7:]
                        if raw_txt.endswith("```"):
                            raw_txt = raw_txt[:-3]
                        parsed_data = json.loads(raw_txt.strip())
                except Exception:
                    try:
                        if free_key_2:
                            genai.configure(api_key=free_key_2)
                            model_b2 = genai.GenerativeModel('gemini-3.6-flash')
                            response = model_b2.generate_content([
                                {'mime_type': file_type, 'data': file_bytes},
                                prompt_text
                            ])
                            raw_txt = response.text.strip()
                            if raw_txt.startswith("```json"):
                                raw_txt = raw_txt[7:]
                            if raw_txt.endswith("```"):
                                raw_txt = raw_txt[:-3]
                            parsed_data = json.loads(raw_txt.strip())
                    except Exception:
                        try:
                            if paid_api_key:
                                genai.configure(api_key=paid_api_key)
                                model_bp = genai.GenerativeModel('gemini-3.6-flash')
                                response = model_bp.generate_content([
                                    {'mime_type': file_type, 'data': file_bytes},
                                    prompt_text
                                ])
                                raw_txt = response.text.strip()
                                if raw_txt.startswith("```json"):
                                    raw_txt = raw_txt[7:]
                                if raw_txt.endswith("```"):
                                    raw_txt = raw_txt[:-3]
                                parsed_data = json.loads(raw_txt.strip())
                        except Exception as batch_err:
                            st.warning(f"No se pudo procesar el archivo {file.name}: {batch_err}")

                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in st.session_state["processed_signatures"] or doc_signature in processed_in_this_batch:
                        duplicate_count += 1
                        st.warning(f"⚠️ Archivo omitido por duplicidad: **{file.name}** (Doc: {num_doc}, Total: {total_doc})")
                    else:
                        processed_in_this_batch.add(doc_signature)
                        st.session_state["processed_signatures"].add(doc_signature)
                        items = parsed_data.get("items", [])
                        if isinstance(items, list):
                            all_consolidated_items.extend(items)

                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.text("¡Procesamiento por lotes completado!")
            
            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y bloquearon **{duplicate_count} archivo(s) duplicado(s)** en este lote.")

            if all_consolidated_items:
                st.success(f"🎉 Se consolidaron exitosamente {len(all_consolidated_items)} ítems de facturas válidas.")

                rows_preview = []
                for idx, item in enumerate(all_consolidated_items, start=1):
                    costo = safe_float(item.get("costo_sin_itbis", 0))
                    raw_pv = (costo * 1.25) * 1.18
                    precio_venta = round_to_nearest_5(raw_pv)
                    cant_comprada = safe_int(item.get("cantidad", 1), 1)
                    empaque_val = safe_int(item.get("empaque", 1), 1)
                    stock_val = cant_comprada * empaque_val
                    codigo_barras = str(item.get("codigo", "")).strip()
                    
                    rows_preview.append({
                        "No.": idx,
                        "Código Barra": codigo_barras,
                        "Nombre": str(item.get("descripcion", "")),
                        "Cant. Compra": cant_comprada,
                        "Empaque": empaque_val,
                        "Stock Total": stock_val,
                        "Costo Unit. Sin ITBIS": costo,
                        "Precio Venta (M5)": precio_venta
                    })

                df_batch = pd.DataFrame(rows_preview)
                st.dataframe(df_batch, use_container_width=True, hide_index=True)

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

                for item_dict in all_consolidated_items:
                    costo = safe_float(item_dict.get("costo_sin_itbis", 0))
                    raw_pv = (costo * 1.25) * 1.18
                    pv = round_to_nearest_5(raw_pv)
                    cant_comprada = safe_int(item_dict.get("cantidad", 1), 1)
                    empaque_val = safe_int(item_dict.get("empaque", 1), 1)
                    stock_val = cant_comprada * empaque_val
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

                output = io.BytesIO()
                wb.save(output)
                excel_data_batch = output.getvalue()

                st.download_button(
                    label="📥 Descargar Excel Consolidado Sin Duplicados",
                    data=excel_data_batch,
                    file_name="Inventario_WilPOS_Consolidado_Lote.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No hay ítems válidos para consolidar (todos los archivos eran duplicados o vacíos).")
