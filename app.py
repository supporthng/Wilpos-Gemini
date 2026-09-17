def process_invoice_exact_18(file_obj, file_type, use_paid_gemini=False, use_openai_fallback=False):
    prompt_text = (
        "Analiza esta factura o tiquet con máxima precisión quirúrgica. "
        "REGLA CRÍTICA DE EXTRACCIÓN LITERAL: Copia exactamente el nombre del producto tal como aparece impreso en la factura en el campo 'descripcion'. "
        "NO inventes, NO adivines ni sustituyas nombres de productos (por ejemplo, si dice 'Parlante', debes escribir 'Parlante' y nunca un licor o marca que no esté escrita). "
        "Para cada renglón extrae estrictamente: "
        "1. 'descripcion': texto literal del producto en la factura. "
        "2. 'cantidad': cantidad numérica. "
        "3. 'unidad': unidad (ej: 'EA', 'CAJA', 'BOT.'). "
        "4. 'tamano': tamaño o presentación si se indica. "
        "5. 'precio_lista': precio unitario o valor de línea. "
        "6. 'valor': monto total neto de la línea. "
        "7. 'descuento_porcentaje': porcentaje de descuento si aplica. "
        "Devuelve un JSON puro bajo la clave 'items': "
        '{"items": [{"descripcion": "Parlante", "cantidad": 1, "unidad": "EA", "tamano": "", "precio_lista": 4063.56, "valor": 4063.56, "descuento_porcentaje": 0.0}]}. '
        "Respuesta JSON pura sin texto adicional."
    )

    active_key = ACTIVE_GEMINI_PAID_KEY if use_paid_gemini else ACTIVE_GEMINI_FREE_KEY
    if not active_key and use_paid_gemini:
        active_key = ACTIVE_GEMINI_FREE_KEY

    for intento in range(2):
        try:
            genai.configure(api_key=active_key)
            model = genai.GenerativeModel('gemini-3.6-flash')
            file_obj.seek(0)
            file_bytes = file_obj.read()
            
            if "pdf" in file_type.lower():
                image_input = {"mime_type": "application/pdf", "data": file_bytes}
            else:
                image_input = Image.open(io.BytesIO(file_bytes))

            response = model.generate_content([image_input, prompt_text])
            raw_text = response.text.strip()
            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]
            return json.loads(raw_text.strip()), "✅ Éxito"
        except Exception as e:
            last_err = str(e)
            if "429" in last_err or "quota" in last_err.lower():
                if not use_paid_gemini and ACTIVE_GEMINI_PAID_KEY:
                    active_key = ACTIVE_GEMINI_PAID_KEY
                    continue
                return None, "QUOTA_EXCEEDED"
            time.sleep(1)
            
    return None, last_err
