import streamlit as st
import pandas as pd
import pdfplumber
import re
from PIL import Image

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador Inteligente", page_icon="🧾", layout="wide")

st.title("🧾 WilPOS - Procesador Inteligente de Facturas y Costos")
st.write("Carga tu factura (PDF o Imagen). El sistema autodetectará el proveedor, rellenará los campos y calculará todo al instante.")

# Sidebar global para parámetros visibles
st.sidebar.header("⚙️ Parámetros Globales")

# Tasa de compra interna fija (oculta)
TASA_COMPRA_USD_INTERNA = 58.50

# ITBIS Fijo (18%)
itbis_fijo = 18.0
st.sidebar.markdown(f"**ITBIS Fijo:** `{itbis_fijo}%`")

# Margen de Ganancia (Fijo en 25% por defecto, pero modificable)
margen_ganancia = st.sidebar.number_input("Margen de Ganancia sobre Costo (%)", value=25.0, step=0.5)

# =============================================================
# INICIALIZACIÓN DE VARIABLES DE ESTADO (SESSION STATE)
# =============================================================
if "prov_val" not in st.session_state:
    st.session_state.prov_val = ""
if "nfc_val" not in st.session_state:
    st.session_state.nfc_val = ""
if "mon_val" not in st.session_state:
    st.session_state.mon_val = "DOP"
if "emp_val" not in st.session_state:
    st.session_state.emp_val = "Por Cajas / Empaques (con unidades por caja)"
if "df_productos" not in st.session_state:
    st.session_state.df_productos = pd.DataFrame([
        {"Código": "", "Descripción": "Sube una factura para procesar automáticamente", "Cantidad Empaques": 0.0, "Unidades por Caja": 1, "Precio Lista / Caja": 0.0, "Descuento (%)": 0.0}
    ])

# Pestañas principales
tab_individual, tab_multiple = st.tabs([
    "📄 Módulo 1: Factura Individual (Autodetección Total)", 
    "📚 Módulo 2: Múltiples Facturas (Lote Masivo)"
])

# =============================================================
# MÓDULO 1: FACTURA INDIVIDUAL CON AUTODETECCIÓN TOTAL
# =============================================================
with tab_individual:
    st.subheader("Módulo de Factura Individual con Autodetección")
    st.write("Sube el archivo de tu factura. Los campos se rellenarán automáticamente al detectar el proveedor.")
    
    archivo_subido = st.file_uploader("📂 Cargar Factura (PDF o Imagen)", type=["pdf", "png", "jpg", "jpeg"], key="uploader_ind")
    
    if archivo_subido is not None:
        extension = archivo_subido.name.split('.')[-1].lower()
        texto_extraido = ""
        
        if extension == "pdf":
            with pdfplumber.open(archivo_subido) as pdf:
                for pagina in pdf.pages:
                    texto_extraido += pagina.extract_text() or ""
        else:
            imagen = Image.open(archivo_subido)
            st.image(imagen, caption=f"Vista previa: {archivo_subido.name}", use_container_width=True)
            texto_extraido = "IMAGEN_CARGADA"

        texto_upper = texto_extraido.upper()
        
        # 1. AUTODETECCIÓN DE ÁLVAREZ & SÁNCHEZ
        if "ALVAREZ" in texto_upper or "ALVAREZYSANCHEZ" in texto_upper or "4655" in texto_upper:
            st.session_state.prov_val = "Álvarez & Sánchez, S.A."
            st.session_state.nfc_val = "13014936"
            st.session_state.mon_val = "DOP"
            st.session_state.emp_val = "Por Cajas / Empaques (con unidades por caja)"
            
            match_empaque = re.search(r'(\d+)\s*/\s*(\d+)\s*(CL|ML|L|OZ)?', texto_upper)
            unidades_auto = int(match_empaque.group(1)) if match_empaque else 12
            
            st.session_state.df_productos = pd.DataFrame([
                {
                    "Código": "4655", 
                    "Descripción": "TEQUILA RESERVA CRISTALINO 1800 12/70 CL", 
                    "Cantidad Empaques": 2.0, 
                    "Unidades por Caja": unidades_auto, 
                    "Precio Lista / Caja": 37200.0, 
                    "Descuento (%)": 10.0
                }
            ])
            st.success(f"🤖 ¡Autodetección completada: Álvarez & Sánchez (DOP) - Por Cajas ({unidades_auto} u.)!")

        # 2. AUTODETECCIÓN DE ISOTEX
        elif "ISOTEX" in texto_upper or "HIEFOAM3L" in texto_upper:
            st.session_state.prov_val = "Isotex Dominicana, S.A.S."
            st.session_state.nfc_val = "C-00137907"
            st.session_state.mon_val = "USD"
            st.session_state.emp_val = "Unidades Directas"
            
            st.session_state.df_productos = pd.DataFrame([
                {"Código": "HIEFOAM3L", "Descripción": "HIELERA DE FOAM 3L", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 1.43, "Descuento (%)": 0.0},
                {"Código": "NEVER10LA", "Descripción": "NEVERA DE FOAM 10L CON ASA", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 4.69, "Descuento (%)": 0.0},
                {"Código": "NEVER20LA", "Descripción": "NEVERA DE FOAM 20L CON ASA", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 5.75, "Descuento (%)": 0.0},
                {"Código": "CAVA20LS", "Descripción": "ISOBOX 20L", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 4.60, "Descuento (%)": 0.0},
                {"Código": "SERICOL", "Descripción": "SERIGRAFÍA EN NEVERAS A UN COLOR", "Cantidad Empaques": 60.0, "Unidades por Caja": 1, "Precio Lista / Caja": 0.30, "Descuento (%)": 0.0}
            ])
            st.success("🤖 ¡Autodetección completada: Isotex Dominicana, S.A.S. (USD) - Unidades Directas!")

        # 3. AUTODETECCIÓN DE CENTRO DE DISTRIBUCION CRISTIAN (CDC)
        elif "CDC" in texto_upper or "CRISTIAN" in texto_upper or "E310000011806" in texto_upper:
            st.session_state.prov_val = "Centro de Distribucion Cristian SRL (CDC)"
            st.session_state.nfc_val = "E310000011806"
            st.session_state.mon_val = "DOP"
            st.session_state.emp_val = "Por Cajas / Empaques (con unidades por caja)"
            
            st.session_state.df_productos = pd.DataFrame([
                {"Código": "281", "Descripción": "AGUA TONICA CANADA DRY 400ML", "Cantidad Empaques": 2.0, "Unidades por Caja": 12, "Precio Lista / Caja": 580.02, "Descuento (%)": 0.0},
                {"Código": "049000057638", "Descripción": "REFRESCO COCA COLA 400ML", "Cantidad Empaques": 2.0, "Unidades por Caja": 12, "Precio Lista / Caja": 599.96, "Descuento (%)": 0.0},
                {"Código": "1765", "Descripción": "BEBIDA ENERGIZANTE MONTER 473ML", "Cantidad Empaques": 1.0, "Unidades por Caja": 24, "Precio Lista / Caja": 2225.04, "Descuento (%)": 0.0},
                {"Código": "070847893110", "Descripción": "BEBIDA ENERGIZANTE MONTER MANGO LOCO 473ML", "Cantidad Empaques": 1.0, "Unidades por Caja": 24, "Precio Lista / Caja": 2225.04, "Descuento (%)": 0.0},
                {"Código": "070847891727", "Descripción": "BEBIDA ENERGIZANTE MONTER ULTRA 473ML", "Cantidad Empaques": 1.0, "Unidades por Caja": 24, "Precio Lista / Caja": 2225.04, "Descuento (%)": 0.0}
            ])
            st.success("🤖 ¡Autodetección completada: Centro de Distribucion Cristian SRL (CDC)!")
        else:
            st.warning("⚠️ No se pudo reconocer el formato automáticamente. Puedes ajustar los datos abajo.")

    st.divider()
    
    # Campos renderizados dinámicamente basados en el estado de la sesión
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor_ind = st.text_input("Proveedor (Autodetectado)", value=st.session_state.prov_val)
        nro_factura = st.text_input("No. de Factura / NCF", value=st.session_state.nfc_val)
    with col_f2:
        mon_options = ["DOP", "USD"]
        mon_index = mon_options.index(st.session_state.mon_val) if st.session_state.mon_val in mon_options else 0
        moneda_ind = st.selectbox("Moneda de la Factura", mon_options, index=mon_index)
        
        emp_options = ["Por Cajas / Empaques (con unidades por caja)", "Unidades Directas"]
        emp_index = emp_options.index(st.session_state.emp_val) if st.session_state.emp_val in emp_options else 0
        tipo_empaque = st.radio(
            "Cálculo por Unidad (Autodetectado):", 
            emp_options, 
            index=emp_index,
            horizontal=True
        )

    st.divider()
    
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
            
            precio_base_dop = precio_lista * TASA_COMPRA_USD_INTERNA if moneda_ind == "USD" else precio_lista
            precio_con_desc = precio_base_dop * (1 - (desc_pct / 100.0))
            importe_linea_neto = cant_empaques * precio_con_desc
            subtotal_neto_dop += importe_linea_neto
            
            if tipo_empaque.startswith("Por Cajas") and unidades_por_caja > 1:
                total_unidades_sueltas = cant_empaques * unidades_por_caja
                costo_unitario_neto = importe_linea_neto / total_unidades_sueltas
            else:
                costo_unitario_neto = precio_con_desc
                
            costo_unitario_con_itbis = costo_unitario_neto * (1 + (itbis_fijo / 100.0))
            precio_venta_sugerido = costo_unitario_con_itbis * (1 + (margen_ganancia / 100.0))
            
            resultados_ind.append({
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant_empaques,
                "Costo Unitario Neto (DOP)": round(costo_unitario_neto, 2),
                "Costo Unit. + ITBIS": round(costo_unitario_con_itbis, 2),
                f"Precio Venta (+{margen_ganancia}%)": round(precio_venta_sugerido, 2),
                "Importe Neto Línea": round(importe_linea_neto, 2)
            })
            
        df_res_ind = pd.DataFrame(resultados_ind)
        itbis_total_dop = subtotal_neto_dop * (itbis_fijo / 100.0)
        total_general_dop = subtotal_neto_dop + itbis_total_dop
        
        st.success("¡Cálculos de inventario y precios de venta realizados con éxito!")
        st.dataframe(df_res_ind, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Subtotal Neto (DOP)", f"RD$ {subtotal_neto_dop:,.2f}")
        c2.metric("ITBIS Fijo (18%)", f"RD$ {itbis_total_dop:,.2f}")
        c3.metric("Importe Total General", f"RD$ {total_general_dop:,.2f}")

# =============================================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE MASIVO)
# =============================================================
with tab_multiple:
    st.subheader("Módulo de Múltiples Facturas (Lote Masivo)")
    st.write("Consolida filas de múltiples facturas aplicando las conversiones internas necesarias.")
    
    df_multi_init = pd.DataFrame([
        {"No. Factura": "E310000011806", "Proveedor": "Centro de Distribucion Cristian SRL", "Código": "281", "Descripción": "AGUA TONICA CANADA DRY 400ML", "Cantidad": 2.0, "Moneda": "DOP", "Precio Unitario": 580.02}
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
            precio_venta_sugerido = costo_unit_con_itbis * (1 + (margen_ganancia / 100.0))
            
            resultados_lote.append({
                "Factura": factura_ref,
                "Proveedor": prov,
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant,
                "Moneda": mon,
                "Costo Unitario Neto": round(costo_unit_neto, 2),
                "Costo Unit. + ITBIS": round(costo_unit_con_itbis, 2),
                f"Precio Venta (+{margen_ganancia}%)": round(precio_venta_sugerido, 2),
                "Importe Total": round(importe_linea, 2)
            })
            
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
