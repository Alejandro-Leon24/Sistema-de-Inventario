# Sistema de Inventario

Aplicacion web para gestion de inventario, ubicaciones y generacion de actas (Entrega, Recepcion y Aula).

## Requisitos

- Python 3.11+ (recomendado 3.12)
- Windows, Linux o macOS
- Dependencias de `requirements.txt`

## Instalacion rapida

1. Crear entorno virtual:

```powershell
python -m venv .venv
```

2. Activar entorno virtual (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

3. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

4. Ejecutar la aplicacion:

```powershell
python app/app.py
```

Aplicacion disponible en `http://127.0.0.1:5000`.

## Base de datos y datos locales

- La base local por defecto es `inventario.db` en la raiz del proyecto.
- La base de datos local **no debe versionarse**.
- El esquema se inicializa/migra automaticamente al arrancar via `database/schema_manager.py`.

## Politica de repositorio (importante)

No subir al repositorio:

- Bases de datos: `inventario.db`, `*.db`, `*.sqlite*`
- Binarios de Python: `__pycache__/`, `*.pyc`
- Salidas de runtime: `salidas/`

Si ya existen archivos trackeados en git, quitarlos del indice sin borrar localmente:

```powershell
git rm --cached inventario.db
git rm --cached -r app/__pycache__ database/__pycache__ app/blueprints/__pycache__ app/utils/__pycache__
```

Luego confirmar con commit.

## Script destructivo de limpieza

`database/update.py` elimina por completo `inventario_items` y reinicia su secuencia.

Ahora requiere todas estas salvaguardas:

- `--confirm-reset`
- `--confirmation-code DELETE-ALL-INVENTORY`
- Variable de entorno `INVENTARIO_ALLOW_DESTRUCTIVE=1`
- Confirmacion interactiva (`RESET`) salvo uso de `--yes`
- Backup automatico en `database/backups/` (excepto si se usa `--skip-backup`)

Ejemplo:

```powershell
$env:INVENTARIO_ALLOW_DESTRUCTIVE = "1"
python database/update.py --confirm-reset --confirmation-code DELETE-ALL-INVENTORY
```

## Pruebas

```powershell
pytest test_api.py -q
```

## Despliegue (guia minima)

- Usar una base de datos fuera del repo.
- Definir carpeta de salida y backups con permisos controlados.
- Respaldar DB antes de actualizar version.
- Ejecutar smoke tests despues de despliegue.
