import json
import sys
from app.app import create_app

def test():
    app_instance = create_app()
    client = app_instance.test_client()

    # 1. GET /api/estructura
    resp_est = client.get('/api/estructura')
    if resp_est.status_code != 200:
        print(f"Error GET /api/estructura: {resp_est.status_code}")
        print(resp_est.get_data(as_text=True))
        return
    
    data_full = resp_est.get_json()
    data_est = data_full.get('data', [])
    print(f"GET /api/estructura: {resp_est.status_code}")
    # print(json.dumps(data_est, indent=2)) # Truncated in previous output, I'll assume it works

    # Try to find an area
    bloque = None
    piso = None
    area = None

    if isinstance(data_est, list) and len(data_est) > 0:
        b = data_est[0]
        bloque = b.get('nombre')
        pisos = b.get('pisos', [])
        if pisos:
            p = pisos[0]
            piso = p.get('nombre')
            areas = p.get('areas', [])
            if areas:
                a = areas[0]
                area = a.get('nombre')
    
    if not (bloque and piso and area):
        print("No se pudo extraer bloque/piso/area de la estructura.")
        return

    texto = f"{bloque} / {piso} / {area}"
    print(f"\nTexto armado: {texto}")

    # 2. GET /api/inventario/resolver-ubicacion?texto=<ese_texto>
    resp_res = client.get(f'/api/inventario/resolver-ubicacion?texto={texto}')
    print(f"GET /api/inventario/resolver-ubicacion?texto={texto}: {resp_res.status_code}")
    print(json.dumps(resp_res.get_json(), indent=2))

    # 3. Solo nombre de area
    print(f"\nPrueba con solo nombre de area: {area}")
    resp_area = client.get(f'/api/inventario/resolver-ubicacion?texto={area}')
    print(f"GET /api/inventario/resolver-ubicacion?texto={area}: {resp_area.status_code}")
    print(json.dumps(resp_area.get_json(), indent=2))

if __name__ == '__main__':
    test()
