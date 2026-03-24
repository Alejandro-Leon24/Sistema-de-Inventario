import json
import urllib.request

# Crear algunos estados de TEST
estados = [
    {"nombre": "Bueno", "descripcion": "En buen estado"},
    {"nombre": "Regular", "descripcion": "En estado regular"},
    {"nombre": "Dañado", "descripcion": "Dañado o inservible"},
]

for estado in estados:
    data = json.dumps(estado).encode('utf-8')
    req = urllib.request.Request('http://localhost:5000/api/parametros/estados', 
                                   data=data, 
                                   headers={'Content-Type': 'application/json'},
                                   method='POST')
    try:
        with urllib.request.urlopen(req) as response:
            print(f"Created: {response.read().decode()}")
    except Exception as e:
        print(f"Error: {e}")

# Crear condiciones
condiciones = [
    {"nombre": "Excelente", "descripcion": "Condición excelente"},
    {"nombre": "Buena", "descripcion": "Condición buena"},
    {"nombre": "Aceptable", "descripcion": "Condición aceptable"},
]

for condicion in condiciones:
    data = json.dumps(condicion).encode('utf-8')
    req = urllib.request.Request('http://localhost:5000/api/parametros/condiciones', 
                                   data=data, 
                                   headers={'Content-Type': 'application/json'},
                                   method='POST')
    try:
        with urllib.request.urlopen(req) as response:
            print(f"Created condition: {response.read().decode()}")
    except Exception as e:
        print(f"Error: {e}")

# Verificar que se crearon
try:
    with urllib.request.urlopen('http://localhost:5000/api/parametros/estados') as response:
        data = json.loads(response.read().decode())
        print("\nEstados creados:")
        print(json.dumps(data, indent=2))
except Exception as e:
    print(f"Error getting states: {e}")

# Verificar condiciones
try:
    with urllib.request.urlopen('http://localhost:5000/api/parametros/condiciones') as response:
        data = json.loads(response.read().decode())
        print("\nCondiciones creadas:")
        print(json.dumps(data, indent=2))
except Exception as e:
    print(f"Error getting conditions: {e}")
