import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="WilPOS - Sistema de Procesamiento", page_icon="🏢", layout="wide")

st.title("🏢 WilPOS - Sistema Modular de Facturación e Inventario")
st.write("Gestiona el procesamiento de costos unitarios e impuestos separando las operaciones locales e internacionales de manera dinámica.")

# Sidebar global para parámetros fiscales
st.sidebar.header("⚙️ Configuración Global")
tasa_usd = st.sidebar.number_input("Tasa de Cambio USD a DOP", value=58.96, step=0.01)
itbis_porcentaje = st.sidebar.slider("Porcentaje de ITBIS (%)", min_value=0.0, max_value=18.0, value=18.0, step=0.5)

# Creación de Módulos Separados por Pestañas
modulo_dop, modulo_usd, modulo_resumen = st.tabs([
    "📥 Módulo 1: Facturas Locales (DOP)", 
    "💱 Módulo 2: Facturas Extranjeras (USD)", 
    "📊 Módulo 3: Consolidado General"
])

# Inicializar estado de sesión para acumular datos entre módulos si es necesario
if "datos_consolidados" not in st.session_state:
    st.session_state.datos_consolidados = pd.DataFrame()

# ==========================================
# MÓDULO 1: FACTURAS LOCALES (DOP)
# ==========================================
with modulo_dop:
    st.subheader("Módulo de Procesamiento - Facturas en Pesos Dominicanos (DOP)")
    st.write("Ingresa o pega los datos de tus proveedores locales (ej. Álvarez & Sánchez u otros). El sistema calculará automáticamente los costos netos e unitarios con descuento e ITBIS.")
    
    df_input_dop = pd.DataFrame([
        {"Código": "4655", "Descripción": "TEQUILA RESERVA CRISTALINO 1800 12/70 CL", "Cantidad": 2.0, "Precio Lista (DOP)": 37200.0, "Descuento (%)": 10.0}
    ])
    
    df_editable_dop = st.data_editor(df_input_dop, num_rows="dynamic", key="editor_dop", use_container_width=True)
    
    if st.button("🚀 Procesar Módulo Local (DOP)", type="primary", key="btn_dop"):
        resultados_dop = []
        subtotal_dop = 0.0
        itbis_dop_total = 0.0
        
        for idx, row in df_editable_dop.iterrows():
            codigo = str(row.get("Código", f"DOP-{idx+1}"))
            desc = str(row.get("Descripción", ""))
            cant = float(row.get("Cantidad", 0.0))
            precio_lista = float(row.get("Precio Lista (DOP)", 0.0))
            desc_pct = float(row.get("Descuento (%)", 0.0))
            
            if cant <= 0 or precio_lista <= 0:
                continue
                
            # Cálculo con descuento y proporción de ITBIS
            precio_con_desc = precio_lista * (1 - (desc_pct / 100.0))
            importe_neto = cant * precio_con_desc
            itbis_linea = importe_neto * (itbis_porcentaje / 100.0)
            
            costo_neto_unitario = precio_con_desc
            costo_unitario_con_itbis = costo_neto_unitario * (1 + (itbis_porcentaje / 100.0))
            
            subtotal_dop += importe_neto
            itbis_dop_total += itbis_linea
            
            resultados_dop.append({
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant,
                "Moneda": "DOP",
                "Costo Unitario Neto": round(costo_neto_unitario, 2),
                "Costo Unitario + ITBIS": round(costo_unitario_con_itbis, 2),
                "Importe Total Neto": round(importe_neto, 2)
            })
            
        df_res_dop = pd.DataFrame(resultados_dop)
        st.success("¡Facturas locales procesadas con éxito!")
        st.dataframe(df_res_dop, use_container_width=True)
        
        # Métricas del módulo
        col1, col2, col3 = st.columns(3)
        col1.metric("Subtotal Neto DOP", f"RD$ {subtotal_dop:,.2f}")
        col2.metric("ITBIS DOP", f"RD$ {itbis_dop_total:,.2f}")
        col3.metric("Total General DOP", f"RD$ {subtotal_dop + itbis_dop_total:,.2f}")

# ==========================================
# MÓDULO 2: FACTURAS EN EXTRANJERO (USD)
# ==========================================
with modulo_usd:
    st.subheader("Módulo de Procesamiento - Cotizaciones / Facturas en Dólares (USD)")
    st.write(f"Los montos ingresados en USD se convertirán automáticamente usando la tasa actual configurada (1 USD = {tasa_usd} DOP).")
    
    df_input_usd = pd.DataFrame([
        {"Código": "HIEFOAM3L", "Descripción": "HIELERA DE FOAM 3L", "Cantidad": 30.0, "Precio Unitario (USD)": 1.43},
        {"Código": "NEVER10LA", "Descripción": "NEVERA DE FOAM 10L CON ASA", "Cantidad": 30.0, "Precio Unitario (USD)": 4.69},
        {"Código": "NEVER20LA", "Descripción": "NEVERA DE FOAM 20L CON ASA", "Cantidad": 30.0, "Precio Unitario (USD)": 5.75},
        {"Código": "CAVA20LS", "Descripción": "ISOBOX 20L", "Cantidad": 30.0, "Precio Unitario (USD)": 4.60},
        {"Código": "SERICOL", "Descripción": "SERIGRAFÍA EN NEVERAS A UN COLOR", "Cantidad": 60.0, "Precio Unitario (USD)": 0.30}
    ])
    
    df_editable_usd = st.data_editor(df_input_usd, num_rows="dynamic", key="editor_usd", use_container_width=True)
    
    if st.button("🚀 Procesar Módulo Internacional (USD)", type="primary", key="btn_usd"):
        resultados_usd = []
        subtotal_usd_val = 0.0
        
        for idx, row in df_editable_usd.iterrows():
            codigo = str(row.get("Código", f"USD-{idx+1}"))
            desc = str(row.get("Descripción", ""))
            cant = float(row.get("Cantidad", 0.0))
            precio_usd = float(row.get("Precio Unitario (USD)", 0.0))
            
            if cant <= 0 or precio_usd <= 0:
                continue
                
            # Conversión a DOP
            precio_dop = precio_usd * tasa_usd
            importe_neto_dop = cant * precio_dop
            
            costo_unitario_neto_dop = precio_dop
            costo_unitario_con_itbis_dop = costo_unitario_neto_dop * (1 + (itbis_porcentaje / 100.0))
            
            subtotal_usd_val += importe_neto_dop
            
            resultados_usd.append({
                "Código": codigo,
                "Descripción": desc,
                "Cantidad": cant,
                "Moneda": "USD->DOP",
                "Costo Unitario Neto (DOP)": round(costo_unitario_neto_dop, 2),
                "Costo Unitario + ITBIS (DOP)": round(costo_unitario_con_itbis_dop, 2),
                "Importe Neto (DOP)": round(importe_neto_dop, 2)
            })
            
        df_res_usd = pd.DataFrame(resultados_usd)
        st.success("¡Documentos internacionales convertidos y procesados con éxito!")
        st.dataframe(df_res_usd, use_container_width=True)
        
        itbis_usd_total = subtotal_usd_val * (itbis_porcentaje / 100.0)
        col1, col2, col3 = st.columns(3)
        col1.metric("Subtotal Convertido DOP", f"RD$ {subtotal_usd_val:,.2f}")
        col2.metric("ITBIS Calculado DOP", f"RD$ {itbis_usd_total:,.2f}")
        col3.metric("Total General DOP", f"RD$ {subtotal_usd_val + itbis_usd_total:,.2f}")

# ==========================================
# MÓDULO 3: CONSOLIDADO GENERAL
# ==========================================
with modulo_resumen:
    st.subheader("📊 Consolidado y Exportación para WilPOS")
    st.write("Aquí puedes visualizar los formatos listos para integrar o descargar en tu sistema central.")
    st.info("Utiliza los botones de procesamiento en los módulos 1 y 2 para calcular los costos unitarios específicos.")
