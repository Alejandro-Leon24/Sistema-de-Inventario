import json
from app.app import create_app

app = create_app()
with app.app_context():
    client = app.test_client()

    textos = [
        "PASILLO SEGUNDO PISO - SOCIOLOGÍA",
        "BLOQUE \"C\" - SOCIOLOGIA / CUARTO PISO ALTO / PASILLO",
        "BLOQUE \"A\" - PRINCIPAL / TERCER PISO ALTO / PASILLO"
    ]

    results = []
    for txt in textos:
        response = client.get(f'/api/inventario/resolver-ubicacion?texto={txt}')
        data = response.get_json()
        
        match_info = "None"
        if data and isinstance(data, dict) and data.get('success'):
            match_data = data.get('match')
            if match_data and match_data.get('area_id'):
                match_info = f"area_id: {match_data['area_id']}, display: {match_data.get('display', 'N/A')}"
        
        results.append({
            'texto': txt,
            'response': data,
            'match': match_info
        })

    print(json.dumps(results, indent=2, ensure_ascii=False))
