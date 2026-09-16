def process_invoice_with_ai(file_obj, file_type, use_openai_fallback=False):
    if use_openai_fallback:
        if not ACTIVE_OPENAI_KEY:
            return None, "Falta clave API de OpenAI"
        import base64
        file_obj.seek(0)
        file_bytes = file_obj.read()
        b64_data = base64.b64encode(file_bytes).decode('utf-8')
        data_url = f"data:application/pdf;base64,{b64_data}" if "pdf" in file_type.lower() else f"data:image/jpeg;base64,{b64_data}"

        prompt_text = (
            "Analiza esta factura COMPLETAMENTE. Extrae TODOS los ítems de la tabla. "
            "Para cada ítem, extrae estrictamente: 'descripcion', 'cantidad', 'empaque', "
            "'precio_lista' (el precio unitario o de caja indicado en la columna PRECIO antes de descuento), "
            "y 'descuento_porcentaje' (el porcentaje de descuento de la columna COM. o DESC., ejemplo: 10 para 10%, o 0). "
            "Devuelve un JSON puro con esta estructura exacta: "
            '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "precio_lista": 0.0, "descuento_porcentaje": 0.0}]}. '
            "Respuesta JSON pura sin texto adicional ni markdown."
        )
        try:
            client = OpenAI(api_key=ACTIVE_OPENAI_KEY)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": [{"type": "text", "text": prompt_text}, {"type": "image_url", "image_url": {"url": data_url}}]}],
                max_tokens=3000
            )
            raw_text = response.choices[0].message.content.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito (OpenAI)"
        except Exception as e:
            return None, str(e)

    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini"

    prompt_text = (
        "Analiza esta factura COMPLETAMENTE. Extrae TODOS los ítems de la tabla. "
        "Para cada ítem, extrae estrictamente: 'descripcion', 'cantidad', 'empaque', "
        "'precio_lista' (el precio unitario o de caja indicado en la columna PRECIO antes de descuento), "
        "y 'descuento_porcentaje' (el porcentaje de descuento de la columna COM. o DESC., ejemplo: 10 para 10%, o 0). "
        "Devuelve un JSON puro con esta estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "precio_lista": 0.0, "descuento_porcentaje": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            image_input = file_bytes if "pdf" in file_type.lower() else Image.open(io.BytesIO(file_bytes))
            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower():
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
    return None, last_err
