import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador General de Facturas", page_icon="⚙️", layout="wide")

st.title("⚙️ WilPOS - Motor de Procesamiento General de Facturas")
st.write("Ingresa o pega los datos de cualquier factura o cotización. El sistema calculará automáticamente los costos unitarios, impuestos y conversiones de moneda.")

# Sidebar para configuraciones globales
st.sidebar.header("Parámetros Fiscales y Financieros")
tasa_usd = st.sidebar.number_input("Tasa de Cambio USD a DOP", value=58.96, step=0.01)
itbis_porcentaje = st.sidebar.slider("Porcentaje de ITBIS (%)", min_value=0.0, max_value=18.0, value=18.0, step=0.5)

st.divider()

# Sección de Entrada Dinámica de Datos
st.subheader("📋 Detalle de la Factura / Cotización")
st.write("Modifica, agrega o pega los ítems del documento a procesar:")

# DataFrame inicial vacío o con una estructura genérica para que el usuario la llene o pegue datos
datos_iniciales = pd.DataFrame([
    {"Código": "PROD-001", "Descripción": "Ejemplo Producto 1", "Cantidad": 10.0, "Precio Unitario": 100.0, "Moneda": "DOP", "Descuento (%)": 0.0},
    {"Código": "PROD-002", "Descripción": "Ejemplo Producto 2 (USD)", "Cantidad": 5.0, "Precio Unitario": 20.0, "Moneda": "USD", "Descuento (%)": 10.0}
])

# st.data_editor permite pegar desde Excel o modificar celdas libremente de forma robusta
df_entrada = st.data_editor(datos_iniciales, num_rows="dynamic", use_container_width=True)

# Lógica General y Robusta de Procesamiento
if st.button("🚀 Procesar e Integrar al Inventario", type="primary"):
    if df_entrada.empty:
        st.warning("Por favor, ingresa al menos un ítem para procesar.")
    else:
        resultados = []
        subtotal_general_dop = 0.0
        itbis_general_dop = 0.0
        
        for index, row in df_entrada.iterrows():
            codigo = str(row.get("Código", f"ITEM-{index+1}"))
            descripcion = str(row.get("Descripción", "Sin descripción"))
            cantidad = float(row.get("Cantidad", 0.0))
            precio_unitario = float(row.get("Precio Unitario", 0.0))
            moneda = str(row.get("Moneda", "DOP")).upper()
            descuento_pct = float(row.get("Descuento (%)", 0.0))
            
            if cantidad <= 0 or precio_unitario <= 0:
                continue
                
            # 1. Normalización de Moneda (Convertir todo a DOP si viene en USD)
            precio_base_dop = precio_unitario * tasa_usd if moneda == "USD" else precio_unitario
            
            # 2. Aplicación de Descuento Comercial por línea
            monto_con_descuento_dop = precio_base_dop * (1 - (descuento_pct / 100.0))
            
            # 3. Cálculo de importes y costos unitarios netos
            importe_linea_neto_dop = cantidad * monto_con_descuento_dop
            
            # Costo unitario neto final en DOP (ya con descuento aplicado pero sin ITBIS)
            costo_unitario_neto_dop = monto_con_descuento_dop
            
            # Costo unitario final incluyendo la proporción del ITBIS
            costo_unitario_con_itbis_dop = costo_unitario_neto_dop * (1 + (itbis_porcentaje / 100.0))
            
            # ITBIS de la línea
            itbis_linea_dop = importe_linea_neto_dop * (itbis_porcentaje / 100.0)
            
            subtotal_general_dop += importe_linea_neto_dop
            itbis_general_dop += itbis_linea_dop
            
            resultados.append({
                "Código": codigo,
                "Descripción": descripcion,
                "Cantidad": cantidad,
                "Costo Unitario Neto (DOP)": round(costo_unitario_neto_dop, 2),
                "Costo Unitario + ITBIS (DOP)": round(costo_unitario_con_itbis_dop, 2),
                "Importe Total Neto (DOP)": round(importe_linea_neto_dop, 2)
            })
            
        df_resultados = pd.DataFrame(resultados)
        total_general_dop = subtotal_general_dop + itbis_general_dop
        
        st.success("¡Documento procesado con éxito mediante la lógica general!")
        
        # Mostrar métricas globales
        col1, col2, col3 = st.columns(3)
        col1.metric("Subtotal General (DOP)", f"RD$ {subtotal_general_dop:,.2f}")
        col2.metric("ITBIS Total (DOP)", f"RD$ {itbis_general_dop:,.2f}")
        col3.metric("Importe Total General (DOP)", f"RD$ {total_general_dop:,.2f}")
        
        # Mostrar tabla formateada lista para exportar o registrar en WilPOS
        st.subheader("📊 Costos Unitarios Calculados para Inventario")
        st.dataframe(df_resultados, use_container_width=True)
        
        # Opción para exportar a CSV para carga masiva
        csv_data = df_resultados.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV de Importación para WilPOS",
            data=csv_data,
            file_name="importacion_wilpos_procesada.csv",
            mime="text/csv"
        )
