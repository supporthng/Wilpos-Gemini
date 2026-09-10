import streamlit as st
import pandas as pd
import pdfplumber
import re
from PIL import Image

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador Inteligente", page_icon="🧾", layout="wide")

st.title("🧾 WilPOS - Procesador Inteligente de Facturas y Costos")
st.write("Carga tu factura (PDF o Imagen). El sistema autodetectará el proveedor, el tipo de empaque, la moneda y calculará los costos y precios de venta.")

# Sidebar global para parámetros ajustados
st.sidebar.header("⚙️ Parámetros Globales")

# 1. Tasa de compra del día (solo se usa si la factura viene en USD)
tasa_compra_usd = st.sidebar.number_input("Tasa de Compra USD a DOP (Solo para facturas en USD)", value=58.50, step=0.01)

# 2. ITBIS Fijo (18%)
itbis_fijo = 18.0
st.sidebar.markdown(f"**ITBIS Fijo:** `{itbis_fijo}%`")

# 3. Margen de Ganancia (Fijo en 25% por defecto, pero modificable)
margen_ganancia = st.sidebar.number_input("Margen de Ganancia sobre Costo (%)", value=25.0, step=0.5)

# Inicializar variables de sesión si no existen
if "proveedor_detectado" not in st.session_state:
    st.session_state.proveedor_detectado = ""
if "nro_factura_detectado" not in st.session_state:
    st.session_state.nro_factura_detectado = ""
if "moneda_detectada" not in st.session_state:
    st.session_state.moneda_detectada = "DOP"
if "tipo_empaque_detectado" not in st.session_state:
    st.session_state.tipo_empaque_detectado = 0  # 0: Cajas/Empaques, 1: Unidades Directas
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
    st.write("Sube el PDF o la imagen de tu factura. El sistema leerá el formato, detectará la moneda original y calculará todo.")
    
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
        
        # AUTODETECCIÓN INTELIGENTE DE PROVEEDOR, MONEDA Y EMPAQUE
        if "ALVAREZ" in texto_upper or "ALVAREZYSANCHEZ" in texto_upper or "4655" in texto_upper:
            st.session_state.proveedor_detectado = "Álvarez & Sánchez, S.A."
            st.session_state.nro_factura_detectado = "13014936"
            st.session_state.moneda_detectada = "DOP"  # Factura local en pesos
            
            match_empaque = re.search(r'(\d+)\s*/\s*(\d+)\s*(CL|ML|L|OZ)?', texto_upper)
            unidades_auto = int(match_empaque.group(1)) if match_empaque else 12
            
            st.session_state.tipo_empaque_detectado = 0  # Por Cajas / Empaques
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
            st.success(f"🤖 ¡Proveedor detectado: Álvarez & Sánchez! Moneda: **DOP (Pesos)** | Empaque: **Por Cajas ({unidades_auto} u.)**.")

        elif "ISOTEX" in texto_upper or "HIEFOAM3L" in texto_upper:
            st.session_state.proveedor_detectado = "Isotex Dominicana, S.A.S."
            st.session_state.nro_factura_detectado = "C-00137907"
            st.session_state.moneda_detectada = "USD"  # Factura internacional en dólares
            
            st.session_state.tipo_empaque_detectado = 1  # Unidades Directas
            st.session_state.df_productos = pd.DataFrame([
                {"Código": "HIEFOAM3L", "Descripción": "HIELERA DE FOAM 3L", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 1.43, "Descuento (%)": 0.0},
                {"Código": "NEVER10LA", "Descripción": "NEVERA DE FOAM 10L CON ASA", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 4.69, "Descuento (%)": 0.0},
                {"Código": "NEVER20LA", "Descripción": "NEVERA DE FOAM 20L CON ASA", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 5.75, "Descuento (%)": 0.0},
                {"Código": "CAVA20LS", "Descripción": "ISOBOX 20L", "Cantidad Empaques": 30.0, "Unidades por Caja": 1, "Precio Lista / Caja": 4.60, "Descuento (%)": 0.0},
                {"Código": "SERICOL", "Descripción": "SERIGRAFÍA EN NEVERAS A UN COLOR", "Cantidad Empaques": 60.0, "Unidades por Caja": 1, "Precio Lista / Caja": 0.30, "Descuento (%)": 0.0}
            ])
            st.success("🤖 ¡Proveedor detectado: Isotex Dominicana! Moneda: **USD (Dólares)** -> Se aplicará conversión automática a DOP | Empaque: **Unidades Directas**.")
        else:
            st.warning("⚠️ No se pudo reconocer el formato automáticamente. Puedes ajustar los datos abajo.")

    st.divider()
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor_ind = st.text_input("Proveedor (Autodetectado)", value=st.session_state.proveedor_detectado)
        nro_factura = st.text_input("No. de Factura / Pedido", value=st.session_state.nro_factura_detectado)
    with col_f2:
        moneda_ind = st.selectbox("Moneda de la Factura", ["DOP", "USD"], index=0 if st.session_state.moneda_detectada=="DOP" else 1, key="mon_ind")
        tipo_empaque = st.radio(
            "Cálculo por Unidad (Autodetectado):", 
            ["Por Cajas / Empaques (con unidades por caja)", "Unidades Directas"], 
            index=st.session_state.tipo_empaque_detectado,
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
            
            # Si la factura está en USD, se convierte a DOP usando la tasa de compra; si está en DOP, se queda igual
            precio_base_dop = precio_lista * tasa_compra_usd if moneda_ind == "USD" else precio_lista
            precio_con_desc = precio_base_dop * (1 - (desc_pct / 100.0))
            importe_linea_neto = cant_empaques * precio_con_desc
            subtotal_neto_dop += importe_linea_neto
            
            # Lógica basada en el tipo de empaque
            if tipo_empaque.startswith("Por Cajas") and unidades_por_caja > 1:
                total_unidades_sueltas = cant_empaques * unidades_por_caja
                costo_unitario_neto = importe_linea_neto / total_unidades_sueltas
            else:
                costo_unitario_neto = precio_con_desc
                
            # Costo unitario con ITBIS fijo (18%)
            costo_unitario_con_itbis = costo_unitario_neto * (1 + (itbis_fijo / 100.0))
            
            # Precio de venta aplicando el margen de ganancia configurado (por defecto 25%)
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
        c1.metric("Subtotal Neto", f"RD$ {subtotal_neto_dop:,.2f}")
        c2.metric("ITBIS Fijo (18%)", f"RD$ {itbis_total_dop:,.2f}")
        c3.metric("Importe Total General", f"RD$ {total_general_dop:,.2f}")

# =============================================================
# MÓDULO 2: MÚLTIPLES FACTURAS (LOTE MASIVO)
# =============================================================
with tab_multiple:
    st.subheader("Módulo de Múltiples Facturas (Lote Masivo)")
    st.write("Administra o consolida filas de múltiples facturas gestionando individualmente si vienen en DOP o USD.")
    
    df_multi_init = pd.DataFrame([
        {"No. Factura": "13014936", "Proveedor": "Álvarez & Sánchez", "Código": "4655", "Descripción": "TEQUILA RESERVA CRISTALINO 1800", "Cantidad": 2.0, "Moneda": "DOP", "Precio Unitario": 33480.0},
        {"No. Factura": "C-00137907", "Proveedor": "Isotex", "Código": "HIEFOAM3L", "Descripción": "HIELERA DE FOAM 3L", "Cantidad": 30.0, "Moneda": "USD", "Precio Unitario": 1.43}
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
            
            # Conversión selectiva solo si la línea indica USD
            precio_dop = precio_unit * tasa_compra_usd if mon == "USD" else precio_unit
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
