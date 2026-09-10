import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="WilPOS - Procesador de Facturas e Inventario", page_icon="🧾", layout="wide")

st.title("🧾 WilPOS - Sistema de Procesamiento de Facturas")
st.write("Gestiona el inventario procesando facturas individuales a detalle o consolidando lotes de múltiples facturas de forma masiva.")

# Sidebar global para parámetros fiscales y de conversión
st.sidebar.header("⚙️ Parámetros Globales")
tasa_usd = st.sidebar.number_input("Tasa de Cambio USD a DOP", value=58.96, step=0.01)
itbis_porcentaje = st.sidebar.slider("Porcentaje de ITBIS (%)", min_value=0.0, max_value=18.0, value=18.0, step=0.5)

# Pestañas principales separadas por el flujo de trabajo correcto
tab_individual, tab_multiple = st.tabs([
    "📄 Módulo 1: Facturas Individuales", 
    "📚 Módulo 2: Múltiples Facturas (Lote Masivo)"
])

# =============================================================
# MÓDULO 1: FACTURAS INDIVIDUALES
# =============================================================
with tab_individual:
    st.subheader("Módulo de Facturas Individuales")
    st.write("Ideal para auditar y procesar el detalle exacto de una factura o cotización individual (cajas, descuentos y costos por botella/unidad).")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor_ind = st.text_input("Proveedor", "Ej. Álvarez & Sánchez, S.A.")
        nro_factura = st.text_input("No. de Factura / Pedido", "Ej. 13014936")
    with col_f2:
        moneda_ind = st.selectbox("Moneda de la Factura", ["DOP", "USD"], key="mon_ind")
        tipo_empaque = st.radio("Cálculo por Unidad:", ["Por Cajas / Empaques (con unidades por caja)", "Unidades Directas"], horizontal=True)

    st.divider()
    
    # Tabla editable para la factura individual
    df_ind_init = pd.DataFrame([
        {
            "Código": "4655", 
            "Descripción": "TEQUILA RESERVA CRISTALINO 1800", 
            "Cantidad Empaques": 2.0, 
            "Unidades por Caja": 12, 
            "Precio Lista / Caja": 37200.0, 
            "Descuento (%)": 10.0
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
            
            # Normalizar precio a DOP si viene en USD
            precio_base_dop = precio_lista * tasa_usd if moneda_ind == "USD" else precio_lista
            
            # Aplicar descuento comercial de la línea
            precio_con_desc = precio_base_dop * (1 - (desc_pct / 100.0))
            importe_linea_neto = cant_empaques * precio_con_desc
            subtotal_neto_dop += importe_linea_neto
            
            # Cálculo de costo unitario final (por botella/unidad suelta)
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
        
        st.success(f"¡Factura '{nro_factura}' procesada con éxito!")
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
    st.write("Agrega, pega o consolida ítems provenientes de varias facturas o proveedores diferentes en una sola tabla general.")
    
    df_multi_init = pd.DataFrame([
        {"No. Factura": "13014936", "Proveedor": "Álvarez & Sánchez", "Código": "4655", "Descripción": "TEQUILA RESERVA CRISTALINO 1800", "Cantidad": 2.0, "Moneda": "DOP", "Precio Unitario": 33480.0},
        {"No. Factura": "C-00137907", "Proveedor": "Isotex", "Código": "HIEFOAM3L", "Descripción": "HIELERA DE FOAM 3L", "Cantidad": 30.0, "Moneda": "USD", "Precio Unitario": 1.43}
    ])
    
    df_multi_edit = st.data_editor(df_multi_init, num_rows="dynamic", key="editor_multiple", use_container_width=True)
    
    if st.button("🚀 Consolidar y Procesar Lote de Múltiples Facturas", type="primary", key="btn_multi"):
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
            
            # Conversión dinámica si la línea está en USD
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
        
        st.success("¡Lote de múltiples facturas consolidado con éxito!")
        st.dataframe(df_res_lote, use_container_width=True)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Subtotal Lote", f"RD$ {subtotal_lote_dop:,.2f}")
        m2.metric("ITBIS Lote", f"RD$ {itbis_lote_dop:,.2f}")
        m3.metric("Total General del Lote", f"RD$ {total_lote_dop:,.2f}")
        
        # Botón para descargar el archivo de importación consolidado
        csv_lote = df_res_lote.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar CSV Consolidado para Importación en WilPOS",
            data=csv_lote,
            file_name="wilpos_lote_multiples_facturas.csv",
            mime="text/csv"
        )
