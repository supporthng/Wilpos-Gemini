import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import io

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador de Facturas e Inventario", page_icon="🧾", layout="wide")

st.title("🧾 WilPOS - Sistema de Procesamiento y Carga de Facturas")
st.write("Carga tus facturas en PDF o en formato de imagen (PNG, JPG), o procesa los datos de manera individual y masiva.")

# Sidebar global para parámetros fiscales y de conversión
st.sidebar.header("⚙️ Parámetros Globales")
tasa_usd = st.sidebar.number_input("Tasa de Cambio USD a DOP", value=58.96, step=0.01)
itbis_porcentaje = st.sidebar.slider("Porcentaje de ITBIS (%)", min_value=0.0, max_value=18.0, value=18.0, step=0.5)

# Pestañas principales
tab_individual, tab_multiple = st.tabs([
    "📄 Módulo 1: Factura Individual (Carga y Detalle)", 
    "📚 Módulo 2: Múltiples Facturas (Lote Masivo)"
])

# =============================================================
# MÓDULO 1: FACTURA INDIVIDUAL (CON SOPORTE PDF E IMÁGENES)
# =============================================================
with tab_individual:
    st.subheader("Módulo de Factura Individual")
    st.write("Sube el archivo de la factura (PDF o imagen PNG/JPG) para auditar y registrar su contenido.")
    
    # Widget de carga actualizado para aceptar PDFs e Imágenes
    archivo_subido = st.file_uploader("📂 Cargar Factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="uploader_ind")
    
    if archivo_subido is not None:
        extension = archivo_subido.name.split('.')[-1].lower()
        
        if extension == "pdf":
            texto_extraido = ""
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    texto_extraido += pagina.extract_text() or ""
            st.success("¡Factura PDF cargada y leída con éxito!")
            with st.expander("🔍 Ver texto bruto extraído del PDF"):
                st.text(texto_extraido)
                
        elif extension in ["png", "jpg", "jpeg"]:
            imagen = Image.open(archivo_subido)
            st.success("¡Imagen de factura cargada con éxito!")
            st.image(imagen, caption=f"Vista previa: {archivo_subido.name}", use_column_width=True)

    st.divider()
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor_ind = st.text_input("Proveedor", "Ej. Álvarez & Sánchez, S.A. / Isotex")
        nro_factura = st.text_input("No. de Factura / Pedido", "Ej. 13014936 o C-00137907")
    with col_f2:
        moneda_ind = st.selectbox("Moneda de la Factura", ["DOP", "USD"], key="mon_ind")
        tipo_empaque = st.radio("Cálculo por Unidad:", ["Por Cajas / Empaques (con unidades por caja)", "Unidades Directas"], horizontal=True)

    st.divider()
    
    df_ind_init = pd.DataFrame([
        {
            "Código": "", 
            "Descripción": "Sube un archivo o escribe los datos aquí", 
            "Cantidad Empaques": 1.0, 
            "Unidades por Caja": 1, 
            "Precio Lista / Caja": 0.0, 
            "Descuento (%)": 0.0
        }
    ])
    
    df_ind_edit = st.data_editor(df_ind_init, num_rows="dynamic", key="editor_individual", use_container_width=True)
    
    if st.button("🧮 Calcular Costos Unitarios de Factura Individual", type="primary", key="btn_ind"):
        subtotal_neto_dop = 0.0
        resultados_ind = []
        
        for idx, row in df_ind_edit.iterrows():
            codigo = str(row.get("Código", f"PROD-{idx+1}"))
            desc = str(row.get("Descripción", ""))
            cant_empaques = float(row.get("Cantidad Empaques", 0.0))
            unidades_por_caja = int(row.get("Unidades por Caja", 1))
            precio_lista = float(row.get("Precio Lista / Caja", 0.0))
            desc_pct = float(row.get("Descuento (%)", 0.0))
            
            if cant_empaques <= 0 or precio_lista <= 0:
                continue
            
            precio_base_dop = precio_lista * tasa_usd if moneda_ind == "USD" else precio_lista
            precio_con_desc = precio_base_dop * (1 - (desc_pct / 100.0))
            importe_linea_neto = cant_empaques * precio_con_desc
            subtotal_neto_dop += importe_linea_neto
            
            if tipo_empaque.startswith("Por Cajas") and unidades_por_caja > 1:
                total_unidades_sueltas = cant_empaques * unidades_por_caja
                costo_unitario_neto = importe_linea_neto / total_unidades_sueltas
            else:
                costo_unitario_neto = precio_con_desc
                
            costo_unitario_con_itbis = costo_unitario_neto * (1 + (itbis_porcentaje / 100.0))
            
            resultados_ind.append({
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant_empaques,
                "Costo Unitario Neto (DOP)": round(costo_unitario_neto, 2),
                "Costo Unitario + ITBIS (DOP)": round(costo_unitario_con_itbis, 2),
                "Importe Neto Línea (DOP)": round(importe_linea_neto, 2)
            })
            
        df_res_ind = pd.DataFrame(resultados_ind)
        itbis_total_dop = subtotal_neto_dop * (itbis_porcentaje / 100.0)
        total_general_dop = subtotal_neto_dop + itbis_total_dop
        
        st.success("¡Factura procesada con éxito!")
        st.dataframe(df_res_ind, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Subtotal Neto", f"RD$ {subtotal_neto_dop:,.2f}")
        c2.metric("ITBIS Total", f"RD$ {itbis_total_dop:,.2f}")
        c3.metric("Importe Total General", f"RD$ {total_general_dop:,.2f}")

# =============================================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE MASIVO)
# =============================================================
with tab_multiple:
    st.subheader("Módulo de Múltiples Facturas (Lote Masivo)")
    st.write("Puedes subir varios archivos (PDF o imágenes) o administrar múltiples filas en la tabla de consolidación.")
    
    archivos_multiples = st.file_uploader("📂 Cargar múltiples archivos de factura", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="uploader_multi")
    if archivos_multiples:
        st.info(f"Se han cargado {len(archivos_multiples)} documentos para el lote.")

    df_multi_init = pd.DataFrame([
        {"No. Factura": "", "Proveedor": "", "Código": "", "Descripción": "", "Cantidad": 0.0, "Moneda": "DOP", "Precio Unitario": 0.0}
    ])
    
    df_multi_edit = st.data_editor(df_multi_init, num_rows="dynamic", key="editor_multiple", use_container_width=True)
    
    if st.button("🚀 Consolidar y Procesar Lote Masivo", type="primary", key="btn_multi"):
        resultados_lote = []
        subtotal_lote_dop = 0.0
        
        for idx, row in df_multi_edit.iterrows():
            factura_ref = str(row.get("No. Factura", ""))
            prov = str(row.get("Proveedor", ""))
            codigo = str(row.get("Código", f"MULT-{idx+1}"))
            desc = str(row.get("Descripción", ""))
            cant = float(row.get("Cantidad", 0.0))
            mon = str(row.get("Moneda", "DOP")).upper()
            precio_unit = float(row.get("Precio Unitario", 0.0))
            
            if cant <= 0 or precio_unit <= 0:
                continue
            
            precio_dop = precio_unit * tasa_usd if mon == "USD" else precio_unit
            importe_linea = cant * precio_dop
            subtotal_lote_dop += importe_linea
            
            costo_unit_neto = precio_dop
            costo_unit_con_itbis = costo_unit_neto * (1 + (itbis_porcentaje / 100.0))
            
            resultados_lote.append({
                "Factura": factura_ref,
                "Proveedor": prov,
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant,
                "Costo Unitario Neto (DOP)": round(costo_unit_neto, 2),
                "Costo Unitario + ITBIS (DOP)": round(costo_unit_con_itbis, 2),
                "Importe Total (DOP)": round(importe_linea, 2)
            })
            
        df_res_lote = pd.DataFrame(resultados_lote)
        itbis_lote_dop = subtotal_lote_dop * (itbis_porcentaje / 100.0)
        total_lote_dop = subtotal_lote_dop + itbis_lote_dop
        
        st.success("¡Lote consolidado con éxito!")
        st.dataframe(df_res_lote, use_container_width=True)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Subtotal Lote", f"RD$ {subtotal_lote_dop:,.2f}")
        m2.metric("ITBIS Lote", f"RD$ {itbis_lote_dop:,.2f}")
        m3.metric("Total General del Lote", f"RD$ {total_lote_dop:,.2f}")
        
        csv_lote = df_res_lote.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV Consolidado para Importación en WilPOS",
            data=csv_lote,
            file_name="wilpos_lote_multiples_facturas.csv",
            mime="text/csv"
        )
