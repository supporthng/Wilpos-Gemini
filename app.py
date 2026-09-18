def online_barcode_lookup_open(barcode_str):
    """Consulta en caché local o realiza una búsqueda web estricta para asegurar el producto correcto."""
    cache = st.session_state["barcode_cache"]
    
    # Si el código ya está en caché, lo retornamos
    if barcode_str in cache:
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        model = genai.GenerativeModel('gemini-3.6-flash')
        
        # Prompt estricto enfocado exclusivamente en el código EAN exacto
        prompt = (
            f"Busca en internet el producto comercial, licor o artículo exacto asociado "
            f"únicamente al código de barras EAN/UPC: '{barcode_str}'. "
            "No inventes ni repitas nombres de otros productos. "
            "Devuelve un JSON puro con el nombre comercial exacto, marca y presentación oficial en mayúsculas: "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si no estás 100% seguro del producto de este código exacto, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO EN INTERNET'}."
        )
        
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        
        data = json.loads(txt.strip())
        desc = data.get("descripcion", "PRODUCTO DESCONOCIDO EN INTERNET")
        
        # Solo guardamos en caché si el resultado es válido y real
        if desc and "DESCONOCIDO" not in desc:
            cache[barcode_str] = desc
            st.session_state["barcode_cache"] = cache
            save_cache(cache)
            return desc, "🌐 (Consultado en Web)"
        else:
            return "PRODUCTO DESCONOCIDO EN INTERNET", "⚠️ (No encontrado en web)"
            
    except Exception as e:
        return "PRODUCTO DESCONOCIDO EN INTERNET", f"⚠️ (Error técnico)"
