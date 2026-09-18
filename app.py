def online_barcode_lookup_open(barcode_str):
    """Consulta en caché local o en internet mediante Gemini utilizando Búsqueda Web en vivo."""
    cache = st.session_state["barcode_cache"]
    if barcode_str in cache:
        return cache[barcode_str], "📥 (Desde Caché Local - Instantáneo)"

    try:
        active_key = ACTIVE_GEMINI_PAID_KEY if ACTIVE_GEMINI_PAID_KEY else ACTIVE_GEMINI_FREE_KEY
        genai.configure(api_key=active_key)
        
        # IMPORTANTE: Configuramos el modelo habilitando la herramienta de búsqueda web oficial
        model = genai.GenerativeModel(
            model_name='gemini-3.6-flash',
            tools=[{"google_search": {}}] # <--- ESTO ACTIVA LA BÚSQUEDA WEB EN STREAMLIT
        )
        
        prompt = (
            f"Busca en internet en tiempo real el producto exacto de bebidas, licores o consumo asociado al código de barras EAN/UPC: '{barcode_str}'. "
            "Devuelve un JSON puro con el nombre comercial exacto, marca y presentación oficial en mayúsculas: "
            '{"descripcion": "NOMBRE DEL PRODUCTO Y PRESENTACION"}. '
            "Si de ninguna manera lo encuentras en la web, devuelve {'descripcion': 'PRODUCTO DESCONOCIDO EN INTERNET'}."
        )
        
        response = model.generate_content(prompt)
        txt = response.text.strip()
        if txt.startswith("```json"): txt = txt[7:]
        if txt.endswith("```"): txt = txt[:-3]
        
        data = json.loads(txt.strip())
        desc = data.get("descripcion", "PRODUCTO DESCONOCIDO EN INTERNET")
        
        # Guardar en caché local para que la próxima vez sea instantáneo
        cache[barcode_str] = desc
        st.session_state["barcode_cache"] = cache
        save_cache(cache)
        
        return desc, "🌐 (Consultado en Internet en Vivo)"
    except Exception as e:
        return f"PRODUCTO DESCONOCIDO EN INTERNET", f"⚠️ (Error: {str(e)})"
