from flask import Flask, request, jsonify

app = Flask(__name__)

# Tasa de cambio de referencia para conversión automática
DEFAULT_USD_RATE = 58.96

@app.route('/api/procesar-cotizacion-isotex', methods=['POST'])
def procesar_cotizacion_isotex():
    """
    Procesa los ítems de la cotización de Isotex Dominicana (C-00137907)[cite: 12],
    convierte los precios de USD a DOP y calcula los costos unitarios finales.
    """
    data = request.get_json()
    tasa_cambio = float(data.get('tasa_cambio', DEFAULT_USD_RATE))
    
    # Detalle de la cotización original en USD[cite: 12]
    items = [
        {"codigo": "HIEFOAM3L", "descripcion": "HIELERA DE FOAM 3L", "cantidad": 30.0, "precio_unitario_usd": 1.43},
        {"codigo": "NEVER10LA", "descripcion": "NEVERA DE FOAM 10L CON ASA", "cantidad": 30.0, "precio_unitario_usd": 4.69},
        {"codigo": "NEVER20LA", "descripcion": "NEVERA DE FOAM 20L CON ASA", "cantidad": 30.0, "precio_unitario_usd": 5.75},
        {"codigo": "CAVA20LS", "descripcion": "ISOBOX 20L", "cantidad": 30.0, "precio_unitario_usd": 4.60},
        {"codigo": "SERICOL", "descripcion": "SERIGRAFÍA EN NEVERAS A UN COLOR", "cantidad": 60.0, "precio_unitario_usd": 0.30}
    ]
    
    items_procesados = []
    subtotal_dop = 0
    
    for item in items:
        # Importe en USD con proporción de ITBIS (18% general aplicado en la cotización)[cite: 12]
        importe_usd = item["cantidad"] * item["precio_unitario_usd"]
        importe_dop_neto = importe_usd * tasa_cambio
        
        # Costo unitario convertido y ajustado a DOP
        costo_unitario_dop = round(item["precio_unitario_usd"] * tasa_cambio * 1.18, 2)
        
        items_procesados.append({
            "codigo": item["codigo"],
            "descripcion": item["descripcion"],
            "cantidad": item["cantidad"],
            "costo_unitario_dop": costo_unitario_dop
        })
        subtotal_dop += importe_dop_neto

    itbis_dop = subtotal_dop * 0.18
    total_dop = subtotal_dop + itbis_dop

    return jsonify({
        "documento": "C-00137907",
        "cliente": "DUSP ROYAL CLUB SRL",
        "tasa_aplicada": tasa_cambio,
        "items": items_procesados,
        "subtotal_dop": round(subtotal_dop, 2),
        "itbis_dop": round(itbis_dop, 2),
        "total_general_dop": round(total_dop, 2)
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
