def online_barcode_lookup_open(barcode_str):
    """Consulta en caché local o en internet mediante Gemini con alta precisión comercial en licores y bebidas."""
    cache = st.session_state["barcode_cache"]
    if barcode_str in cache:
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-3.6-flash')
        
        # Prompt optimizado para búsqueda web precisa en bebidas, licores y productos comerciales
        prompt = (
            f"Actúa como un sistema experto de inventario y punto de venta. Busca en la web el producto exacto "
            f"asociado al código de barras EAN/UPC: '{barcode_str}'. "
            "Enfócate prioritariamente en bebidas alcohólicas, licores, vinos, whiskies, tequilas, ginebras o productos de consumo masivo si aplica. "
            "Devuelve un JSON puro con el nombre comercial exacto, marca y presentación oficial en mayúsculas (ej: TITO'S HANDMADE VODKA 50 ML o GIN HENDRICK'S 50 CL): "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si de ninguna manera lo encuentras, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO EN INTERNET'}."
        )
        
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        data = json.loads(txt.strip())
        desc = data.get("descripcion", "PRODUCTO DESCONOCIDO EN INTERNET")
        
        # Guardar automáticamente en caché local para que la próxima vez sea instantáneo
        cache[barcode_str] = desc
        st.session_state["barcode_cache"] = cache
        save_cache(cache)
        
        return desc, "🌐 (Consultado en Internet)"
    except Exception:
        return "PRODUCTO DESCONOCIDO EN INTERNET", "⚠️ (Error de red)"
