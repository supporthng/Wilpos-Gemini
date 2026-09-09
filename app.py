import os
import time
from google import genai
from google.genai import types

# Configura tu API Key de Gemini
# Puedes definir la variable de entorno GEMINI_API_KEY o colocarla directamente aquí
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

CARPETA_BUZON = "./Facturas_Nuevas"
EXCEL_WILPOS = "Inventario_WilPOS_Todas_Facturas_Actualizado.xlsx"

def procesar_facturas_automaticas():
    print("=== MONITOR DE FACTURAS WILPOS ACTIVO ===")
    if not os.path.exists(CARPETA_BUZON):
        os.makedirs(CARPETA_BUZON)
        print(f"Se creó la carpeta '{CARPETA_BUZON}'. Coloca tus facturas ahí.")

    while True:
        archivos = os.listdir(CARPETA_BUZON)
        for archivo in archivos:
            ruta_archivo = os.path.join(CARPETA_BUZON, archivo)
            if os.path.isfile(ruta_archivo) and archivo.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.xml')):
                print(f"\n[!] Nueva factura detectada: {archivo}")
                
                # Subir archivo y consultar a Gemini usando el modelo multimodal
                with open(ruta_archivo, "rb") as f:
                    file_bytes = f.read()
                
                print("Enviando a Gemini para cálculo automático (Sin ITBIS + 25% Venta)...")
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[
                        types.Part.from_bytes(
                            data=file_bytes,
                            mime_type='application/pdf' if archivo.endswith('.pdf') else 'image/jpeg',
                        ),
                        "Lee esta factura, extrae todos los productos, calcula el costo unitario sin el ITBIS (descontando impuestos y descuentos comerciales si aplica), y calcula el precio de venta (costo + 25%). Devuélveme el resultado estructurado para agregarlo a WilPOS."
                    ]
                )
                
                print("\n--- RESPUESTA DE ANÁLISIS ---")
                print(response.text)
                print("-----------------------------\n")
                
                # Mover archivo procesado a una subcarpeta para no repetirlo
                procesados_dir = os.path.join(CARPETA_BUZON, "Procesadas")
                if not os.path.exists(procesados_dir):
                    os.makedirs(procesados_dir)
                os.rename(ruta_archivo, os.path.join(procesados_dir, archivo))
                print(f"[✔] Factura {archivo} procesada y archivada con éxito.")
                
        time.sleep(10) # Revisa la carpeta cada 10 segundos

if __name__ == "__main__":
    procesar_facturas_automaticas()
