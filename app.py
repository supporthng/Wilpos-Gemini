def online_barcode_lookup_open(barcode_str):
    """Consulta segura con base local de licores conocidos y respaldo inteligente."""
    cache = st.session_state["barcode_cache"]
    
    # 1. Base fija directa para tus productos principales (Cero errores de red o IA)
    base_conocida = {
        "5010106113493": "BALLANTINE'S FINEST BLENDED SCOTCH WHISKY 70 CL",
        "082184001363": "TEQUILA PATRÓN REPOSADO 750 ML",
        "5010327709000": "WHISKY GRANT'S TRIPLE WOOD 750 ML"
    }
    
    if barcode_str in base_conocida:
        return base_conocida[barcode_str], "⚡ (Catálogo Rápido WilPOS)"

    # 2. Revisar caché local de productos previamente escaneados y guardados
    if barcode_str in cache and "DESCONOCIDO" not in str(cache[barcode_str]):
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = (
            f"Proporciona el nombre comercial exacto, marca y presentación en mayúsculas "
            f"del licor o bebida asociada al código de barras EAN/UPC: '{barcode_str}'. "
            "Responde únicamente con el nombre (ej: BUCHANAN'S DELUXE 12 AÑOS 750 ML). "
            "Si no lo conoces con absoluta seguridad, responde exactamente: DESCONOCIDO"
        )
        
        response = model.generate_content(prompt)
        desc = response.text.strip().upper()
        
        if desc and "DESCONOCIDO" not in desc and len(desc) > 3:
            cache[barcode_str] = desc
            st.session_state["barcode_cache"] = cache
            save_cache(cache)
            return desc, "🌐 (Consultado con éxito)"
        else:
            return f"ARTICULO NUEVO ({barcode_str})", "✏️ (Asigna el nombre manualmente)"
            
    except Exception:
        return f"ARTICULO NUEVO ({barcode_str})", "✏️ (Asigna el nombre manualmente)"
