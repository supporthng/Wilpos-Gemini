import io
import json
import os
import hashlib
import google.generativeai as genai
from PIL import Image

@@ -31,6 +32,10 @@ def safe_float(val, default=0.0):
def round_to_nearest_5(x):
return round(round(x / 5) * 5, 2)

# Inicializar control de documentos procesados para evitar duplicados
if "processed_signatures" not in st.session_state:
    st.session_state["processed_signatures"] = set()

# ==========================================
# MENÚ DE NAVEGACIÓN LATERAL
# ==========================================
@@ -40,14 +45,19 @@ def round_to_nearest_5(x):
["📄 Factura Individual", "📂 Múltiples Facturas (Lote)"]
)

# Botón para limpiar historial de duplicados si el usuario lo desea
if st.sidebar.button("🧹 Restablecer Memoria de Duplicados"):
    st.session_state["processed_signatures"] = set()
    st.sidebar.success("¡Memoria de duplicados reiniciada!")

# ==========================================
# MÓDULO 1: FACTURA INDIVIDUAL (EL QUE YA TENÍAS)
# MÓDULO 1: FACTURA INDIVIDUAL
# ==========================================
if modulo == "📄 Factura Individual":
st.title("📊 Automatizador de Facturas para WilPOS (Individual)")
    st.markdown("Sube tu factura. La IA detectará los empaques, calculará el costo unitario sin ITBIS y el **stock total correcto**, aplicando la fórmula de WilPOS (**25% margen + 18% ITBIS** con redondeo a **múltiplos de 5**).")
    st.markdown("Sube tu factura. El sistema cuenta con **control anti-duplicados** para evitar procesar dos veces el mismo documento basándose en su emisor, número de factura, fecha y monto total.")

    uploaded_file = st.file_uploader("Sube tu factura (Imagen o PDF)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")
    uploaded_file = st.file_uploader("Sube tu factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="single_file")

if "quota_exceeded" not in st.session_state:
st.session_state["quota_exceeded"] = False
@@ -73,206 +83,89 @@ def paid_confirmation_dialog():
st.rerun()

if uploaded_file is not None:
        st.success(f"¡Factura cargada: {uploaded_file.name}!")
        st.success(f"¡Archivo cargado: {uploaded_file.name}[cite: 12, 13]!")

if st.session_state["quota_exceeded"]:
paid_confirmation_dialog()

        if st.session_state["use_paid_now"]:
            with st.spinner("Procesando factura con la versión de pago..."):
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
                    genai.configure(api_key=paid_api_key)
                    model_paid = genai.GenerativeModel('gemini-3.6-flash')
                    
                    uploaded_file.seek(0)
                    file_bytes = uploaded_file.read()
                    
                    prompt = (
                        "Analiza esta factura detalladamente. Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (la cantidad comprada de cajas/paquetes/unidades, ej: 2, 4, 5, 6), 'empaque' (cuántas unidades individuales trae cada caja o paquete, ej: 12, 24, 10, o 1 si es suelto), y el valor total e ITBIS. "
                        "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'cantidad', 'empaque', 'costo_sin_itbis'). "
                        "REGLA DE ORO PARA EL COSTO UNITARIO: Calcula el valor neto sin ITBIS (Valor Total - ITBIS) y divídelo entre el total de unidades individuales (cantidad * empaque) para obtener el 'costo_sin_itbis' por unidad exacta. "
                        "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML'). "
                        "REGLA CRÍTICA PARA CÓDIGOS DE BARRAS: Extrae rigurosamente el código de barras completo de cada producto (EAN-13, UPC o código de proveedor). Trata el campo 'codigo' estrictamente como texto (string), PRESERVANDO ABSOLUTAMENTE TODOS LOS CEROS A LA IZQUIERDA. "
                        "Respuesta JSON válida sin texto adicional."
                    )
                    
                    response = model_paid.generate_content([
                        {'mime_type': uploaded_file.type if hasattr(uploaded_file, 'type') else 'image/jpeg', 'data': file_bytes},
                        prompt
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
                    raw_text = raw_text.strip()
                    
                    data_items = json.loads(raw_text)
                    st.session_state["use_paid_now"] = False
                    st.success("✅ ¡Factura procesada exitosamente con la versión de pago!")
                    
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
                except Exception as err_paid:
                    st.error(f"Error al procesar con la versión de pago: {err_paid}")
                    st.session_state["use_paid_now"] = False
        else:
            if st.button("🚀 Procesar Factura con Plantilla Oficial"):
                with st.spinner("Analizando empaques, cantidades y calculando stock correcto..."):
                    data_items = None
                    success_msg = ""
                    
                    prompt_text = (
                        "Analiza esta factura detalladamente. Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (la cantidad comprada de cajas/paquetes/unidades, ej: 2, 4, 5, 6), 'empaque' (cuántas unidades individuales trae cada caja o paquete, ej: 12, 24, 10, o 1 si es suelto), y el valor total e ITBIS. "
                        "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'cantidad', 'empaque', 'costo_sin_itbis'). "
                        "REGLA DE ORO PARA EL COSTO UNITARIO: Calcula el valor neto sin ITBIS (Valor Total - ITBIS) y divídelo entre el total de unidades individuales (cantidad * empaque) para obtener el 'costo_sin_itbis' por unidad exacta. "
                        "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML'). "
                        "REGLA CRÍTICA PARA CÓDIGOS DE BARRAS: Extrae rigurosamente el código de barras completo de cada producto (EAN-13, UPC o código de proveedor). Trata el campo 'codigo' estrictamente como texto (string), PRESERVANDO ABSOLUTAMENTE TODOS LOS CEROS A LA IZQUIERDA. "
                        "Respuesta JSON válida sin texto adicional."
                    )
                    
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
                                success_msg = "✅ ¡Factura procesada con éxito rotando al **Respaldo Gratuito #2**!"
                            except Exception as e2:
                                err_msg2 = str(e2)
                                if "429" in err_msg2 or "Quota exceeded" in err_msg2:
                                    st.session_state["quota_exceeded"] = True
                                    st.rerun()
                                else:
                                    st.error(f"Error con Respaldo Gratuito #2: {e2}")
                        elif "429" in err_msg1 or "Quota exceeded" in err_msg1:
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
                            st.error(f"Ocurrió un error con la IA: {e1}")
                    else:
                        st.error(f"Error al procesar: {e}")

                if parsed_data:
                    rnc_emisor = str(parsed_data.get("emisor_rnc", "")).strip()
                    num_doc = str(parsed_data.get("numero_documento", "")).strip()
                    fecha_doc = str(parsed_data.get("fecha", "")).strip()
                    total_doc = str(parsed_data.get("total", "")).strip()
                    
                    signature_string = f"{rnc_emisor}_{num_doc}_{fecha_doc}_{total_doc}"
                    doc_signature = hashlib.md5(signature_string.encode('utf-8')).hexdigest()

                    if data_items:
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
@@ -354,57 +247,54 @@ def paid_confirmation_dialog():
)

# ==========================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE / BATCH)
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE) CON ANTI-DUPLICADOS
# ==========================================
elif modulo == "📂 Múltiples Facturas (Lote)":
    st.title("📂 Procesador por Lotes de Facturas (Múltiples Archivos)")
    st.markdown("Sube varias facturas (PDF o imágenes). El sistema procesará cada una, consolidará los productos en una sola lista y calculará el inventario completo para WilPOS.")
    st.title("📂 Procesador por Lotes con Control Anti-Duplicados")
    st.markdown("Sube varias facturas o cotizaciones. El sistema filtrará automáticamente cualquier documento duplicado evaluando su firma única.")

uploaded_files = st.file_uploader("Sube tus facturas (Puedes seleccionar varias)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="batch_files")

if uploaded_files:
        st.info(f"Se han cargado {len(uploaded_files)} archivos para procesar.")
        st.info(f"Se han cargado {len(uploaded_files)} archivos en total.")

        if st.button("🚀 Procesar Lote Completo de Facturas", type="primary"):
        if st.button("🚀 Procesar Lote Evaluando Duplicados", type="primary"):
all_consolidated_items = []
            duplicate_count = 0
            processed_in_this_batch = set()

progress_bar = st.progress(0)
status_text = st.empty()

for i, file in enumerate(uploaded_files):
                status_text.text(f"Procesando archivo {i+1} de {len(uploaded_files)}: {file.name}...")
                status_text.text(f"Analizando archivo {i+1} de {len(uploaded_files)}: {file.name}[cite: 12, 13]...")

prompt_text = (
                    "Analiza esta factura detalladamente. Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (la cantidad comprada de cajas/paquetes/unidades, ej: 2, 4, 5, 6), 'empaque' (cuántas unidades individuales trae cada caja o paquete, ej: 12, 24, 10, o 1 si es suelto), y el valor total e ITBIS. "
                    "Devuelve la información en formato JSON puro (una lista de objetos con claves exactas: 'codigo', 'descripcion', 'cantidad', 'empaque', 'costo_sin_itbis'). "
                    "REGLA DE ORO PARA EL COSTO UNITARIO: Calcula el valor neto sin ITBIS (Valor Total - ITBIS) y divídelo entre el total de unidades individuales (cantidad * empaque) para obtener el 'costo_sin_itbis' por unidad exacta. "
                    "REGLA PARA DESCRIPCIÓN: Limpia la descripción para que solo incluya el nombre principal y su tamaño (ejemplo: 'BEBIDA ENERGIZANTE CICLON 250ML'). "
                    "REGLA CRÍTICA PARA CÓDIGOS DE BARRAS: Extrae rigurosamente el código de barras completo de cada producto (EAN-13, UPC o código de proveedor). Trata el campo 'codigo' estrictamente como texto (string), PRESERVANDO ABSOLUTAMENTE TODOS LOS CEROS A LA IZQUIERDA. "
                    "Respuesta JSON válida sin texto adicional."
                    "Analiza esta factura o cotización detalladamente. Extrae los datos de cabecera: 'emisor_rnc', 'numero_documento', 'fecha', 'total'. "
                    "Para cada ítem, extrae: 'codigo', 'descripcion', 'cantidad' (cantidad comprada), 'empaque' (unidades por empaque), y 'costo_sin_itbis'. "
                    "Devuelve la información estrictamente en formato JSON con la siguiente estructura exacta: "
                    '{"emisor_rnc": "...", "numero_documento": "...", "fecha": "...", "total": "...", "items": [{"codigo": "...", "descripcion": "...", "cantidad": 1, "empaque": 1, "costo_sin_itbis": 0.0}]}. '
                    "REGLA CRÍTICA: Preserva todos los ceros a la izquierda como texto. Respuesta JSON pura sin texto adicional."
)

file_bytes = file.read()
file_type = file.type if hasattr(file, 'type') else 'image/jpeg'

                items_from_file = None
                
                # Intentar con Cuenta Gratuita 1
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
                        items_from_file = json.loads(raw_txt.strip())
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
                    # Intentar con Cuenta Gratuita 2 en caso de fallo o cuota
try:
if free_key_2:
genai.configure(api_key=free_key_2)
@@ -418,35 +308,38 @@ def paid_confirmation_dialog():
raw_txt = raw_txt[7:]
if raw_txt.endswith("```"):
raw_txt = raw_txt[:-3]
                            items_from_file = json.loads(raw_txt.strip())
                            parsed_data = json.loads(raw_txt.strip())
except Exception:
                        # Fallback final a versión de pago si las gratuitas fallan
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
                                items_from_file = json.loads(raw_txt.strip())
                        except Exception as batch_err:
                            st.warning(f"No se pudo procesar el archivo {file.name}: {batch_err}")
                        pass

                if items_from_file and isinstance(items_from_file, list):
                    all_consolidated_items.extend(items_from_file)
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

            status_text.text("¡Procesamiento por lotes completado con éxito!")
            status_text.text("¡Procesamiento por lotes completado!")
            
            if duplicate_count > 0:
                st.error(f"🚨 Se detectaron y bloquearon **{duplicate_count} archivo(s) duplicado(s)** en este lote.")

if all_consolidated_items:
                st.success(f"🎉 Se consolidaron un total de {len(all_consolidated_items)} ítems provenientes de todas las facturas.")
                st.success(f"🎉 Se consolidaron exitosamente {len(all_consolidated_items)} ítems de facturas válidas.")

rows_preview = []
for idx, item in enumerate(all_consolidated_items, start=1):
@@ -472,7 +365,6 @@ def paid_confirmation_dialog():
df_batch = pd.DataFrame(rows_preview)
st.dataframe(df_batch, use_container_width=True, hide_index=True)

                # Generar Excel Consolidado
template_path = "Plantilla_Inventario_WilPOS_2.xlsx"
if not os.path.exists(template_path):
template_path = "Plantilla_Inventario_WilPOS.xlsx"
@@ -523,10 +415,10 @@ def paid_confirmation_dialog():
excel_data_batch = output.getvalue()

st.download_button(
                    label="📥 Descargar Excel Consolidado (Lote WilPOS)",
                    label="📥 Descargar Excel Consolidado Sin Duplicados",
data=excel_data_batch,
file_name="Inventario_WilPOS_Consolidado_Lote.xlsx",
mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
else:
                st.warning("No se pudieron extraer ítems de las facturas proporcionadas.")
                st.warning("No hay ítems válidos para consolidar (todos los archivos eran duplicados o vacíos).")
