if st.session_state["batch_processed_count"] > 0:
            st.markdown("---")
            st.markdown("## 📊 Consolidado de Inventario y Totales Generales")

            raw_items = st.session_state["batch_accumulated_items"]
            multiplicador_ganancia = 1 + (margen_ganancia_lote / 100.0)

            processed_rows = []
            total_unidades_inventario = 0
            
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                
                desc = str(item.get("descripcion") or item.get("nombre") or item.get("articulo") or item.get("item") or "").strip()
                if not desc:
                    continue

                official_code, matched_name, _ = match_official_barcode(desc)

                raw_costo = safe_float(item.get("costo_sin_itbis") or item.get("costo") or item.get("precio") or 0)
                cant_comprada = safe_int(item.get("cantidad") or item.get("cant") or 1, 1)
                empaque_val = safe_int(item.get("empaque") or item.get("unidad_empaque") or 1, 1)

                costo, empaque_val = audit_and_correct_cost(desc, raw_costo, cant_comprada, empaque_val)
                raw_pv = (costo * multiplicador_ganancia) * 1.18
                precio_venta = round_to_nearest_5(raw_pv)
                stock_val = cant_comprada * empaque_val
                
                total_unidades_inventario += stock_val

                processed_rows.append({
                    "Nombre": matched_name,
                    "Código Barra": str(official_code),
                    "Categoría": "General",
                    "Tipo": "producto",
                    "Precio Venta": precio_venta,
                    "Costo": costo,
                    "Stock": stock_val,
                    "Stock Mínimo": 5,
                    "ITBIS": 0.18,
                    "Unidad Medida": "unidad",
                    "Venta Granel": "No",
                    "Cantidad Empaque": empaque_val,
                    "Precio Variable": "No",
                    "Descuento %": 0,
                    "Descuento Monto": 0,
                    "Precio Especial": None,
                    "Descuento Activo": "No",
                    "Descuento Nota": None
                })

            if processed_rows:
                df_temp = pd.DataFrame(processed_rows)

                if 'Código Barra' not in df_temp.columns:
                    df_temp['Código Barra'] = 'S/C (Sin Código)'
                if 'Nombre' not in df_temp.columns:
                    df_temp['Nombre'] = 'PRODUCTO DESCONOCIDO'

                df_grouped = df_temp.groupby(['Código Barra', 'Nombre'], as_index=False).agg({
                    'Stock': 'sum',
                    'Costo': 'mean',
                    'Precio Venta': 'mean',
                    'Categoría': 'first',
                    'Tipo': 'first',
                    'Stock Mínimo': 'first',
                    'ITBIS': 'first',
                    'Unidad Medida': 'first',
                    'Venta Granel': 'first',
                    'Cantidad Empaque': 'first',
                    'Precio Variable': 'first',
                    'Descuento %': 'first',
                    'Descuento Monto': 'first',
                    'Precio Especial': 'first',
                    'Descuento Activo': 'first',
                    'Descuento Nota': 'first'
                })

                df_final_preview = df_grouped.sort_values(by="Stock", ascending=False).reset_index(drop=True)
                df_final_preview["Código Barra"] = df_final_preview["Código Barra"].astype(str)

                # ==========================================
                # MÉTRICAS CLAVE EN PANTALLA (KPIs)
                # ==========================================
                kpi1, kpi2, kpi3 = st.columns(3)
                kpi1.metric("📁 Facturas Procesadas Exitosamente", f"{st.session_state['batch_ok_count']}")
                kpi2.metric("📦 Total Artículos / Unidades en Stock", f"{total_unidades_inventario:,}")
                
                # Cálculo estimado del costo total de inversión del lote
                inversion_total_lote = (df_final_preview['Costo'] * df_final_preview['Stock']).sum()
                kpi3.metric("💰 Inversión Neta Total (Sin ITBIS)", f"RD$ {inversion_total_lote:,.2f}")

                st.markdown("---")
                st.success(f"✨ Se consolidaron **{len(df_final_preview)} productos únicos** provenientes de tus facturas.")
                st.dataframe(df_final_preview, use_container_width=True, hide_index=True)
