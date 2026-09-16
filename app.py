def process_invoice_with_gemini(file_obj, file_type):
    if not ACTIVE_GEMINI_KEY:
        return None, "Falta clave API de Gemini (Configura GEMINI_API_KEY en st.secrets)"

    prompt_text = (
        "Analiza esta factura con precisión milimétrica. "
        "Extrae cada renglón de la tabla con: "
        "1. 'descripcion': texto exacto de la columna 'DESCRIPCION'. "
        "2. 'cantidad': número de la columna 'CANTIDAD'. "
        "3. 'unidad': 'CAJA' o 'BOT.'. "
        "4. 'tamano': texto exacto de la columna 'TAMAÑO'. "
        "5. 'precio_lista': número exacto de la columna 'PRECIO'. "
        "6. 'descuento_porcentaje': porcentaje exacto de la columna 'COM.'. "
        "Devuelve un JSON puro bajo la clave 'items': "
        '{"items": [{"descripcion": "...", "cantidad": 1, "unidad": "CAJA", "tamano": "12/75 CL.", "precio_lista": 0.0, "descuento_porcentaje": 10.0}]}. '
        "Respuesta JSON pura sin texto adicional ni markdown."
    )

    file_obj.seek(0)
    file_bytes = file_obj.read()
    image_input = Image.open(io.BytesIO(file_bytes)) if "pdf" not in file_type.lower() else file_bytes

    for intento in range(3):
        try:
            if NEW_GOOGLE_SDK:
                client = genai.Client(api_key=ACTIVE_GEMINI_KEY)
                # Usando el cliente moderno con el modelo Flash estándar
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[image_input, prompt_text]
                )
                raw_text = response.text.strip()
            else:
                genai.configure(api_key=ACTIVE_GEMINI_KEY)
                model = genai.GenerativeModel('gemini-2.5-flash')
                response = model.generate_content([image_input, prompt_text])
                raw_text = response.text.strip()

            if not raw_text:
                raise ValueError("Respuesta vacía de la API.")

            if raw_text.startswith("```json"): raw_text = raw_text[7:]
            if raw_text.endswith("```"): raw_text = raw_text[:-3]

            time.sleep(3.5)
            return json.loads(raw_text.strip()), "✅ Éxito (Gemini Flash)"
        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "quota" in err_str.lower() or "RESOURCE_EXHAUSTED" in err_str:
                if intento < 2:
                    run_visual_countdown(25, f"Límite de peticiones alcanzado (Intento {intento+1}/3). Esperando reintento automático:")
                    continue
                return None, "QUOTA_EXCEEDED"
            return None, f"Error técnico API: {err_str}"
    return None, "Error desconocido en API"
