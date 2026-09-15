# ==========================================
# MOTOR DE ASIGNACIÓN OFICIAL POR NOMBRE (SOBREESCRIBIENDO CUALQUIER CÓDIGO DE FACTURA)
# ==========================================
def match_official_barcode(item_description):
    raw_name = str(item_description).strip().upper()
    b_mem = st.session_state["barcode_memory"]
    
    # 1. Coincidencia Exacta
    if raw_name in b_mem:
        return clean_barcode(b_mem[raw_name]), raw_name, "Maestro Exacto"

    # 2. Coincidencia Normalizada Avanzada (Maneja abreviaturas de proveedores)
    norm_memory = get_normalized_memory_dict()
    norm_input = normalize_text(raw_name)
    
    if norm_input in norm_memory:
        code, orig_name = norm_memory[norm_input]
        return clean_barcode(code), orig_name, "Normalizado Avanzado"

    # 3. Coincidencia Fuzzy (Similitud alta >= 0.80)
    best_ratio = 0.0
    best_code = "S/C (Sin Código)"
    best_name = raw_name
    
    for n_key, (code, orig_name) in norm_memory.items():
        ratio = difflib.SequenceMatcher(None, norm_input, n_key).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_code = code
            best_name = orig_name

    if best_ratio >= 0.80:
        return clean_barcode(best_code), best_name, f"Fuzzy ({best_ratio:.2f})"

    return "S/C (Sin Código)", raw_name, "⚠️ Sin Coincidencia en Maestro"
