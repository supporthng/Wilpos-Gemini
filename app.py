def process_invoice_with_ai(file_obj, file_type, use_openai_fallback=False):
    if use_openai_fallback:
        return process_with_openai(file_obj, file_type)

    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini"

    prompt_text = (
        "Analiza esta factura COMPLETAMENTE de arriba a abajo. Extrae TODOS los ítems sin omitir ninguno. "
        "Para cada ítem, extrae estrictamente: 'descripcion', 'cantidad', 'empaque', 'precio_lista' (el precio unitario o de caja antes de descuento indicado en la factura), "
        "y 'descuento_porcentaje' (ejemplo: 10 si tiene 10%, o 0 si no tiene). "
        "Devuelve un JSON puro con esta estructura exacta: "
        '{"emisor_rnc": "...", "emisor_nombre": "...", "numero_documento": "...", "fecha": "...", "total": 0.0, "items": [{"descripcion": "...", "cantidad": 1, "empaque": 1, "precio_lista": 0.0, "descuento_porcentaje": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    last_err = ""
    for intento in range(2):
        try:
            genai.configure(api_key=ACTIVE_GEMINI_KEY)
            model = genai.GenerativeModel('gemini-3.6-flash')
            
            file_obj.seek(0)
            file_bytes = file_obj.read()
            
            if "pdf" in file_type.lower():
                image_input = file_bytes
            else:
                image_input = Image.open(io.BytesIO(file_bytes))

            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
                
            parsed_data = json.loads(raw_text.strip())
            return parsed_data, "✅ Éxito"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower():
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
            continue
    return None, last_err
