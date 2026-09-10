import streamlit as st
import pandas as pd
import pdfplumber
import re
from PIL import Image

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador Inteligente de Facturas", page_icon="🧾", layout="wide")

st.title("🧾 WilPOS - Procesador General y Automático de Facturas")
st.write("Sube cualquier factura en PDF o Imagen. El sistema extraerá las líneas de productos automáticamente de forma genérica.")

# Sidebar global para parámetros visibles
st.sidebar.header("⚙️ Parámetros Globales")

# Tasa de compra interna fija (oculta para conversiones de USD a DOP)
TASA_COMPRA_USD_INTERNA = 58.50

# ITBIS Fijo (18%)
itbis_fijo = 18.0
st.sidebar.markdown(f"**ITBIS Fijo:** `{itbis_fijo}%`")

# Margen de Ganancia (Fijo en 25% por defecto, modificable)
margen_ganancia = st.sidebar.number_input("Margen de Ganancia sobre Costo (%)", value=25.0, step=0.5)

# =============================================================
# INICIALIZACIÓN DE VARIABLES DE ESTADO (SESSION STATE)
# =============================================================
if "prov_val" not in st.session_state:
    st.session_state.prov_val = "Proveedor Genérico"
if "nfc_val" not in st.session_state:
    st.session_state.nfc_val = "AUTODECT"
if "mon_val" not in st.session_state:
    st.session_state.mon_val = "DOP"
if "emp_val" not in st.session_state:
    st.session_state.emp_val = "Por Cajas / Empaques (con unidades por caja)"
if "df_productos" not in st.session_state:
    st.session_state.df_productos = pd.DataFrame([
        {"Código": "", "Descripción": "Sube una factura para extraer ítems automáticamente", "Cantidad Empaques": 1.0, "Unidades por Caja": 1, "Precio Lista / Caja": 0.0, "Descuento (%)": 0.0}
    ])

# Pestañas principales
tab_individual, tab_multiple = st.tabs([
    "📄 Módulo 1: Factura Individual (Extracción Genérica)", 
    "📚 Módulo 2: Múltiples Facturas (Lote Masivo)"
])

# =============================================================
# MÓDULO 1: EXTRACCIÓN GENÉRICA DE CUALQUIER FACTURA
# =============================================================
with tab_individual:
    st.subheader("Módulo de Procesamiento Genérico")
    st.write("Sube cualquier factura PDF. El motor intentará extraer las tablas y descripciones de manera completamente automática.")
    
    archivo_subido = st.file_uploader("📂 Cargar Factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="uploader_ind")
    
    if archivo_subido is not None:
        extension = archivo_subido.name.split('.')[-1].lower()
        texto_extraido = ""
        filas_extraidas = []
        
        if extension == "pdf":
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    # Extraer texto general para detectar moneda o proveedor de forma genérica
                    texto_extraido += pagina.extract_text() or ""
                    
                    # Extraer tablas de manera genérica
                    tablas = pagina.extract_tables()
                    for tabla in tablas:
                        for fila in tabla:
                            # Limpiar celdas vacías o nulas de la fila
                            fila_limpia = [str(c).strip() for c in fila if c is not None and str(c).strip() != ""]
                            if len(fila_limpia >= 2):
                                filas_extraidas.append(fila_limpia)
        else:
            imagen = Image.open(archivo_subido)
            st.image(imagen, caption=f"Vista previa: {archivo_subido.name}", use_container_width=True)
            texto_extraido = "IMAGEN_CARGADA"

        texto_upper = texto_extraido.upper()
        
        # 1. Autodetección genérica de Moneda
        if "USD" in texto_upper or "US$" in texto_upper:
            st.session_state.mon_val = "USD"
        else:
            st.session_state.mon_val = "DOP"
            
        # 2. Autodetección genérica de empaques o porciones (ej: 12/70, Paquete-12, etc.)
        match_empaque = re.search(r'(\d+)\s*(?:/|PAQUETE-|CAJA-)\s*(\d+)?', texto_upper)
        if match_empaque:
            st.session_state.emp_val = "Por Cajas / Empaques (con unidades por caja)"
        
        # Si se extrajeron filas de tablas genéricas, intentamos pasarlas al editor
        if filas_extraidas:
            nuevos_items = []
            for idx, f in enumerate(filas_extraidas):
                # Intentar mapear celdas de forma heurística genérica
                desc = f[1] if len(f) > 1 else "Item extraído"
                if "DESCRIPCION" in desc.upper() or "TOTAL" in desc.upper():
                    continue # Saltar encabezados o totales
                nuevos_items.append({
                    "Código": f[0] if len(f) > 0 else f"GEN-{idx}",
                    "Descripción": desc,
                    "Cantidad Empaques": 1.0,
                    "Unidades por Caja": 1,
                    "Precio Lista / Caja": 0.0,
                    "Descuento (%)": 0.0
                })
            if nuevos_items:
                st.session_state.df_productos = pd.DataFrame(nuevos_items)
                st.success("🤖 ¡Estructura de factura analizada y extraída de manera genérica con éxito!")
        else:
            st.info("ℹ️ Factura leída. Si no se autocompletaron las líneas por diseño gráfico del PDF, puedes ingresarlas o pegarlas abajo.")

        st.rerun()

    st.divider()
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor_ind = st.text_input("Proveedor", value=st.session_state.prov_val)
        nro_factura = st.text_input("No. de Factura / NCF", value=st.session_state.nfc_val)
    with col_f2:
        mon_options = ["DOP", "USD"]
        mon_index = mon_options.index(st.session_state.mon_val) if st.session_state.mon_val in mon_options else 0
        moneda_ind = st.selectbox("Moneda de la Factura", mon_options, index=mon_index)
        
        emp_options = ["Por Cajas / Empaques (con unidades por caja)", "Unidades Directas"]
        emp_index = emp_options.index(st.session_state.emp_val) if st.session_state.emp_val in emp_options else 0
        tipo_empaque = st.radio(
            "Cálculo por Unidad:", 
            emp_options, 
            index=emp_index,
            horizontal=True
        )

    st.divider()
    
    st.write("📋 **Detalle de Ítems (Editable o Pegado Directo)**")
    df_ind_edit = st.data_editor(st.session_state.df_productos, num_rows="dynamic", key="editor_individual", use_container_width=True)
    
    if st.button("🧮 Calcular Costos y Precios de Venta", type="primary", key="btn_ind"):
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
            
            # Conversión interna automática si la moneda es USD
            precio_base_dop = precio_lista * TASA_COMPRA_USD_INTERNA if moneda_ind == "USD" else precio_lista
            precio_con_desc = precio_base_dop * (1 - (desc_pct / 100.0))
            importe_linea_neto = cant_empaques * precio_con_desc
            subtotal_neto_dop += importe_linea_neto
            
            # Costo unitario dinámico según empaque
            if tipo_empaque.startswith("Por Cajas") and unidades_por_caja > 1:
                total_unidades_sueltas = cant_empaques * unidades_por_caja
                costo_unitario_neto = importe_linea_neto / total_unidades_sueltas
            else:
                costo_unitario_neto = precio_con_desc
                
            costo_unitario_con_itbis = costo_unitario_neto * (1 + (itbis_fijo / 100.0))
            
            # Precio de venta aplicando el margen y redondeando al múltiplo de 5
            precio_venta_bruto = costo_unitario_con_itbis * (1 + (margen_ganancia / 100.0))
            precio_venta_sugerido = round(precio_venta_bruto / 5) * 5
            
            resultados_ind.append({
                "Código": codigo,
                "Descripción": desc,
                "Cantidad Empaques": cant_empaques,
                "Costo Unitario Neto (DOP)": round(costo_unitario_neto, 2),
                "Costo Unit. + ITBIS": round(costo_unitario_con_itbis, 2),
                f"Precio Venta (+{margen_ganancia}% - Múltiplo de 5)": round(precio_venta_sugerido, 2),
                "Importe Neto Línea": round(importe_linea_neto, 2)
            })
            
        if resultados_ind:
            df_res_ind = pd.DataFrame(resultados_ind)
            itbis_total_dop = subtotal_neto_dop * (itbis_fijo / 100.0)
            total_general_dop = subtotal_neto_dop + itbis_total_dop
            
            st.success("¡Cálculos de inventario y precios de venta realizados con éxito!")
            st.dataframe(df_res_ind, use_container_width=True)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Subtotal Neto (DOP)", f"RD$ {subtotal_neto_dop:,.2f}")
            c2.metric("ITBIS Fijo (18%)", f"RD$ {itbis_total_dop:,.2f}")
            c3.metric("Importe Total General", f"RD$ {total_general_dop:,.2f}")
        else:
            st.warning("Verifica que los valores de cantidad y precios sean mayores a cero.")

# =============================================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE MASIVO)
# =============================================================
with tab_multiple:
    st.subheader("Módulo de Múltiples Facturas (Lote Masivo)")
    st.write("Consolida filas provenientes de varias facturas de cualquier proveedor.")
    
    df_multi_init = pd.DataFrame([
        {"No. Factura": "FACT-001", "Proveedor": "Proveedor Genérico", "Código": "PROD-01", "Descripción": "Artículo de prueba", "Cantidad": 1.0, "Moneda": "DOP", "Precio Unitario": 100.0}
    ])
    
    df_multi_edit = st.data_editor(df_multi_init, num_rows="dynamic", key="editor_multiple", use_container_width=True)
    
    if st.button("🚀 Consolidar Lote Masivo y Precios", type="primary", key="btn_multi"):
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
            
            precio_dop = precio_unit * TASA_COMPRA_USD_INTERNA if mon == "USD" else precio_unit
            importe_linea = cant * precio_dop
            subtotal_lote_dop += importe_linea
            
            costo_unit_neto = precio_dop
            costo_unit_con_itbis = costo_unit_neto * (1 + (itbis_fijo / 100.0))
            
            precio_venta_bruto = costo_unit_con_itbis * (1 + (margen_ganancia / 100.0))
            precio_venta_sugerido = round(precio_venta_bruto / 5) * 5
            
            resultados_lote.append({
                "Factura": factura_ref,
                "Proveedor": prov,
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant,
                "Moneda": mon,
                "Costo Unitario Neto": round(costo_unit_neto, 2),
                "Costo Unit. + ITBIS": round(costo_unit_con_itbis, 2),
                f"Precio Venta (+{margen_ganancia}% - Múltiplo de 5)": round(precio_venta_sugerido, 2),
                "Importe Total": round(importe_linea, 2)
            })
            
        if resultados_lote:
            df_res_lote = pd.DataFrame(resultados_lote)
            itbis_lote_dop = subtotal_lote_dop * (itbis_fijo / 100.0)
            total_lote_dop = subtotal_lote_dop + itbis_lote_dop
            
            st.success("¡Lote consolidado con éxito!")
            st.dataframe(df_res_lote, use_container_width=True)
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Subtotal Lote", f"RD$ {subtotal_lote_dop:,.2f}")
            m2.metric("ITBIS Fijo (18%)", f"RD$ {itbis_lote_dop:,.2f}")
            m3.metric("Total General del Lote", f"RD$ {total_lote_dop:,.2f}")
            
            csv_lote = df_res_lote.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar CSV Consolidado para Importación en WilPOS",
                data=csv_lote,
                file_name="wilpos_lote_multiples_facturas.csv",
                mime="text/csv"
            )
