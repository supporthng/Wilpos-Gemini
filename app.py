import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import os
import hashlib
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Inteligente de Facturas", page_icon="📊", layout="wide")

# ==========================================
# CONFIGURACIÓN DE CREDENCIALES Y AYUDAS
# ==========================================import streamlit as st
import openpyxl
import pandas as pd
import io
import json
import os
import hashlib
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="WilPOS - Automatizador Inteligente de Facturas", page_icon="📊", layout="wide")

# ==========================================
# CONFIGURACIÓN DE CREDENCIALES Y AYUDAS
# ==========================================
free_key_1 = st.secrets.get("GEMINI_API_KEY", "")
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", "")
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", "")

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

if "processed_signatures" not in st.session_state:
    st.session_state["processed_signatures"] = set()

# ==========================================
# MENÚ DE NAVEGACIÓN LATERAL
# ==========================================
st.sidebar.title("📌 Menú de Módulos")
modulo = st.sidebar.radio(
    "Selecciona el modo de trabajo:",
    ["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

if st.sidebar.button("🧹 Restablecer Memoria de Duplicados"):
    st.session_state["processed_signatures"] = set()
    st.sidebar.success("¡Memoria de duplicados reiniciada!")

# ==========================================
# FUNCIÓN DE CONSOLIDACIÓN CON RASTEO DE FACTURAS
# ==========================================
def aggregate_items_with_sources(items_with_sources_list):
    """
    Agrupa los ítems por código de barras y rastrea exactamente en cuáles 
    facturas/documentos se encuentra cada producto para notificar al usuario.
    """
    aggregated_dict = {}
    
    for entry in items_with_sources_list:
        item = entry["item"]
        source_name = entry["source_name"]
        
        codigo = str(item.get("codigo", "")).strip()
        if not codigo:
            codigo = "SIN_CODIGO"
            
        descripcion = str(item.get("descripcion", "")).strip()
        cant = safe_int(item.get("cantidad", 1), 1)
        empaque = safe_int(item.get("empaque", 1), 1)
        costo = safe_float(item.get("costo_sin_itbis", 0.0))
        
        if codigo in aggregated_dict:
            existing = aggregated_dict[codigo]
            existing["cantidad"] += cant
            existing["costo_sin_itbis"] = (existing["costo_sin_itbis"] + costo) / 2
            if source_name not in existing["sources"]:
                existing["sources"].append(source_name)
        else:
            aggregated_dict[codigo] = {
                "codigo": codigo,
                "descripcion": descripcion,
                "cantidad": cant,
                "empaque": empaque,
                "costo_sin_itbis": costo,
                "sources": [source_name]
            }
            
    return list(aggregated_dict.values())

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
    st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura. El sistema cuenta con control anti-duplicados de documentos.")

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if "quota_exceeded" not in st.session_state:
        st.session_state["quota_exceeded"] = False

    if "use_paid_now" not in st.session_state:
        st.session_state["use_paid_now"] = False

    @st.dialog("⚠️ Límite de Cuota Gratuita Alcanzado")
    def paid_confirmation_dialog():
        st.write("Se han agotado las solicitudes gratuitas de **ambas cuentas** de Gemini.")
        st.write("¿Deseas procesar esta factura utilizando la **versión de pago**?")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Sí, usar Versión de Pago", type="primary", use_container_width=True):
                st.session_state["quota_exceeded"] = False
                st.session_state["use_paid_now"] = True
                st.rerun()
        with col2:
            if st.button("❌ Cancelar", use_container_width=True):
                st.session_state["quota_exceeded"] = False
                st.session_state["use_paid_now"] = False
                st.rerun()

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}!")
        
        if st.session_state["quota_exceeded"]:
            paid_confirmation_dialog()
            
        if st.button("🚀 Procesar Factura") or st.session_state["use_paid_now"]:
            with st.spinner("Analizando factura y verificando duplicidad..."):
                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada), 'empaque' (unidades por empaque), y 'costo_sin_itbis'. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA PARA CÓDIGOS DE BARRAS: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )
                
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                file_type = uploaded_file.type if hasattr(uploaded_file, 'type') else 'application/pdf'
                
                parsed_data = None
                success_msg = ""
                
                active_key = paid_api_key if st.session_state["use_paid_now"] else (free_key_1 if free_key_1 else free_key_2)
                keys_sequence = [active_key]
                if free_key_2 and free_key_2 not in keys_sequence:
                    keys_sequence.append(free_key_2)
                if paid_api_key and paid_api_key not in keys_sequence:
                    keys_sequence.append(paid_api_key)

                for k in keys_sequence:
                    if not k:
                        continue
                    try:
                        genai.configure(api_key=k)
                        model = genai.GenerativeModel('gemini-3.6-flash')
                        response = model.generate_content([
                            {'mime_type': file_type, 'data': file_bytes},
                            prompt_text
                        ])
                        raw_text = response.text.strip()
                        if raw_text.startswith("```json"):
                            raw_text = raw_text[7:]
                        if raw_text.endswith("```"):
                            raw_text = raw_text[:-3]
                        parsed_data = json.loads(raw_text.strip())
                        if parsed_data and isinstance(parsed_data, dict):
                            success_msg = "✅ ¡Factura procesada con éxito!"
                            break
                    except Exception as e:
                        if "429" in str(e) or "Quota exceeded" in str(e):
                            continue
                        else:
                            st.error(f"Error al procesar: {e}")
                            break

                if not parsed_data and not st.session_state["use_paid_now"]:
                    st.session_state["quota_exceeded"] = True
                    st.rerun()

                if parsed_data:
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in st.session_state["processed_signatures"]:
                        st.error("🚨 **¡ADVERTENCIA DE FACTURA DUPLICADA!**")
                        st.warning(f"Este documento ya fue procesado anteriormente.\n\n"
                                   f"- **Emisor RNC:** `{rnc_emisor}`\n"
                                   f"- **No. Documento:** `{num_doc}`\n"
                                   f"- **Fecha:** `{fecha_doc}`\n"
                                   f"- **Total:** `{total_doc}`\n\n"
                                   f"El sistema ha bloqueado la carga para evitar duplicidad.")
                    else:
                        st.session_state["processed_signatures"].add(doc_signature)
                        st.success(success_msg)
                        
                        raw_items = [{"item": it, "source_name": f"{uploaded_file.name} (Doc: {num_doc})"} for it in parsed_data.get("items", [])]
                        data_items = aggregate_items_with_sources(raw_items)
                        
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
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE) CON NOTIFICACIÓN DE PRODUCTOS Y FACTURAS DE ORIGEN
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.title("📂 Procesador por Lotes (Consolidación Inteligente)")
    st.markdown("Sube varias facturas. El sistema consolidará automáticamente los productos repetidos en una sola fila y **te notificará en qué facturas se encuentra cada producto**[cite: 13].")

    uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

    if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")
        
        if st.button("🚀 Procesar Lote y Consolidar Stock", type="primary"):
            all_items_with_sources = []
            duplicate_count = 0
            error_count = 0
            processed_in_this_batch = set()
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, file in enumerate(uploaded_files):
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}...")
                
                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada), 'empaque' (unidades por empaque), y 'costo_sin_itbis'. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )
                
                file.seek(0)
                file_bytes = file.read()
                file_type = file.type if hasattr(file, 'type') else 'application/pdf'
                
                parsed_data = None
                keys_to_try = [k for k in [free_key_1, free_key_2, paid_api_key] if k]
                
                for k in keys_to_try:
                    try:
                        genai.configure(api_key=k)
                        model_b = genai.GenerativeModel('gemini-3.6-flash')
                        response = model_b.generate_content([
                            {'mime_type': file_type, 'data': file_bytes},
                            prompt_text
                        ])
                        raw_txt = response.text.strip()
                        if raw_txt.startswith("```json"):
                            raw_txt = raw_txt[7:]
                        if raw_txt.endswith("```"):
                            raw_txt = raw_txt[:-3]
                        parsed_data = json.loads(raw_txt.strip())
                        if parsed_data and isinstance(parsed_data, dict):
                            break
                    except Exception:
                        continue
                
                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in st.session_state["processed_signatures"] or doc_signature in processed_in_this_batch:
                        duplicate_count += 1
                        st.warning(f"⚠️ Archivo omitido por duplicidad de factura: **{file.name}** (Doc: {num_doc}, Total: {total_doc})")
                    else:
                        processed_in_this_batch.add(doc_signature)
                        st.session_state["processed_signatures"].add(doc_signature)
                        
                        items = parsed_data.get("items", [])
                        source_label = f"{file.name} (Doc: {num_doc})" if num_doc else file.name
                        if isinstance(items, list):
                            for it in items:
                                all_items_with_sources.append({
                                    "item": it,
                                    "source_name": source_label
                                })
                else:
                    error_count += 1
                    st.error(f"❌ No se pudo analizar el archivo: **{file.name}**.")
                
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            status_text.text("¡Procesamiento por lotes completado!")
            
            if duplicate_count > 0:
                st.warning(f"🚨 Se detectaron y bloquearon **{duplicate_count} factura(s) duplicada(s)**.")
            if error_count > 0:
                st.error(f"⚠️ **{error_count} archivo(s)** no pudieron ser procesados.")
            
            # CONSOLIDACIÓN INTELIGENTE CON RASTEO DE ORIGEN
            consolidated_items = aggregate_items_with_sources(all_items_with_sources)
            
            if consolidated_items:
                st.success(f"🎉 Se consolidaron {len(consolidated_items)} productos únicos sin duplicar stock.")
                
                # SECCIÓN DE NOTIFICACIÓN DETALLADA DE PRODUCTOS Y FACTURAS
                with st.expander("📋 Ver detalle de consolidación y facturas de origen de cada producto", expanded=True):
                    for prod in consolidated_items:
                        sources_str = ", ".join([f"**{src}**" for src in prod["sources"]])
                        st.markdown(f"- 📦 **[{prod['codigo']}] {prod['descripcion']}** (Cant. Total Comprada: `{prod['cantidad']}`)\n  - 📄 Encontrado en las facturas: {sources_str}")
                
                rows_preview = []
                for idx, item in enumerate(consolidated_items, start=1):
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
                
                for item_dict in consolidated_items:
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
                    label="📥 Descargar Excel Consolidado Sin Duplicar Productos ni Stock",
                    data=excel_data_batch,
                    file_name="Inventario_WilPOS_Consolidado_Lote.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No hay ítems válidos para consolidar.")
free_key_1 = st.secrets.get("GEMINI_API_KEY", "")
free_key_2 = st.secrets.get("GEMINI_API_KEY_2", "")
paid_api_key = st.secrets.get("GEMINI_API_KEY_PAID", "")

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

# Inicializar control de documentos procesados para evitar duplicados
if "processed_signatures" not in st.session_state:
    st.session_state["processed_signatures"] = set()

# ==========================================
# MENÚ DE NAVEGACIÓN LATERAL
# ==========================================
st.sidebar.title("📌 Menú de Módulos")
modulo = st.sidebar.radio(
    "Selecciona el modo de trabajo:",
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
    st.markdown("Sube tu factura. El sistema cuenta con **control anti-duplicados** para evitar procesar dos veces el mismo documento basándose en su emisor, número de factura, fecha y monto total.")

    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

    if "quota_exceeded" not in st.session_state:
        st.session_state["quota_exceeded"] = False

    if "use_paid_now" not in st.session_state:
        st.session_state["use_paid_now"] = False

    @st.dialog("⚠️ Límite de Cuota Gratuita Alcanzado")
    def paid_confirmation_dialog():
        st.write("Se han agotado las solicitudes gratuitas de **ambas cuentas** de Gemini.")
        st.write("¿Deseas procesar esta factura utilizando la **versión de pago**?")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Sí, usar Versión de Pago", type="primary", use_container_width=True):
                st.session_state["quota_exceeded"] = False
                st.session_state["use_paid_now"] = True
                st.rerun()
        with col2:
            if st.button("❌ Cancelar", use_container_width=True):
                st.session_state["quota_exceeded"] = False
                st.session_state["use_paid_now"] = False
                st.rerun()

    if uploaded_file is not None:
        st.success(f"¡Archivo cargado: {uploaded_file.name}[cite: 12, 13]!")
        
        if st.session_state["quota_exceeded"]:
            paid_confirmation_dialog()
            
        if st.button("🚀 Procesar Factura con Control Anti-Duplicados") or st.session_state["use_paid_now"]:
            with st.spinner("Analizando factura y verificando duplicidad..."):
                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada, ej: 2, 4, 30), 'empaque' (unidades por empaque, ej: 1, 10, 12, 24), y 'costo_sin_itbis' (calculado dividiendo el valor neto sin ITBIS entre el total de unidades individuales: cantidad * empaque). "
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
                    genai.configure(api_key=active_key)
                    model = genai.GenerativeModel('gemini-3.6-flash')
                    response = model.generate_content([
                        {'mime_type': file_type, 'data': file_bytes},
                        prompt_text
                    ])
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    parsed_data = json.loads(raw_text.strip())
                    success_msg = "✅ ¡Factura procesada con éxito!"
                except Exception as e:
                    err_str = str(e)
                    if ("429" in err_str or "Quota exceeded" in err_str) and not st.session_state["use_paid_now"]:
                        try:
                            genai.configure(api_key=free_key_2)
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
                        st.warning(f"Este documento ya fue procesado anteriormente.\n\n"
                                   f"- **Emisor RNC:** `{rnc_emisor}`\n"
                                   f"- **No. Documento:** `{num_doc}`[cite: 12, 13]\n"
                                   f"- **Fecha:** `{fecha_doc}`[cite: 12, 13]\n"
                                   f"- **Total:** `{total_doc}`[cite: 12, 13]\n\n"
                                   f"El sistema ha bloqueado la carga para evitar duplicidad en el inventario de WilPOS.")
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
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE) CON ANTI-DUPLICADOS
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
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}[cite: 12, 13]...")
                
                prompt_text = (
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada), 'empaque' (unidades por empaque), y 'costo_sin_itbis'. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
                )
                
                file_bytes = file.read()
                file_type = file.type if hasattr(file, 'type') else 'image/jpeg'
                
                parsed_data = None
                try:
                    genai.configure(api_key=free_key_1 if free_key_1 else paid_api_key)
                    model_b = genai.GenerativeModel('gemini-3.6-flash')
                    response = model_b.generate_content([
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
                        pass
                
                if parsed_data and isinstance(parsed_data, dict):
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()
                    
                    if doc_signature in st.session_state["processed_signatures"] or doc_signature in processed_in_this_batch:
                        duplicate_count += 1
                        st.warning(f"⚠️ Archivo omitido por duplicidad: **{file.name}**[cite: 12, 13] (Doc: {num_doc}[cite: 12, 13], Total: {total_doc}[cite: 12, 13])")
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
