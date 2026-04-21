# Sistema de Inventario

Aplicación web para gestión de inventario, ubicaciones y generación de actas (Entrega, Recepción, Bajas, Traspaso e Informe por Áreas).

## Características

- Gestión jerárquica de ubicaciones (Bloque > Piso > Área).
- Control detallado de bienes con soporte para códigos ESBYE.
- Importación inteligente desde Excel con mapeo dinámico de columnas.
- Generación de actas en formato **Microsoft Word (DOCX)**.
- Exportación de listados a **Microsoft Excel (XLSX)**.
- Vista previa en vivo de actas mediante renderizado HTML.
- Historial completo de actas generadas con trazabilidad de plantillas.
- Auditoría de cambios en el inventario.

## Requisitos

- Python 3.11+ (recomendado 3.12)
- Windows, Linux o macOS
- Dependencias de `requirements.txt`

## Instalación rápida

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

4. Ejecutar la aplicación:

```powershell
python app/app.py
```

Aplicación disponible en `http://127.0.0.1:5000`.

## Base de datos

- El sistema utiliza **SQLite** con modo WAL para alta concurrencia.
- La base local por defecto es `inventario.db` en la raíz del proyecto.
- El esquema se inicializa/migra automáticamente al arrancar.

## Política de repositorio

No subir al repositorio:
- Archivos de base de datos (`*.db`, `*.sqlite*`).
- Salidas de documentos (`salidas/`).
- Entornos virtuales (`.venv/`).
- Cachés (`__pycache__/`, `.pytest_cache/`).

## Pruebas

```powershell
pytest test_api.py -q
```
