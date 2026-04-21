import os
import sys
from pathlib import Path

# Manejo de rutas para entorno empaquetado (PyInstaller)
if getattr(sys, 'frozen', False):
    # Si la app está empaquetada, BASE_DIR es donde está el .exe
    # sys._MEIPASS es la carpeta temporal de PyInstaller
    BASE_DIR = Path(sys.executable).parent
    BUNDLE_DIR = Path(sys._MEIPASS)
else:
    # Si corre como script normal
    BASE_DIR = Path(__file__).resolve().parent.parent
    BUNDLE_DIR = BASE_DIR

if str(BASE_DIR) in sys.path:
    sys.path.remove(str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR))

from flask import Flask, render_template

def create_app() -> Flask:
    # Configuramos Flask para buscar templates y static en BUNDLE_DIR (interno al exe)
    app = Flask(__name__, 
                template_folder=str(BUNDLE_DIR / "app" / "templates"),
                static_folder=str(BUNDLE_DIR / "app" / "static"))
    
    # La base de datos debe estar en BASE_DIR para que sea persistente (fuera del exe)
    app.config["DATABASE"] = BASE_DIR / "inventario.db"
    
    from database.db import init_app
    init_app(app)

    with app.app_context():
        from database.schema_manager import init_schema
        # Usamos BUNDLE_DIR para leer schema.sql pero BASE_DIR para la DB
        init_schema(BUNDLE_DIR)

    try:
        from app.blueprints import register_blueprints
    except ModuleNotFoundError:
        from blueprints import register_blueprints
    
    register_blueprints(app)

    # Keep legacy endpoint names used by templates after blueprint split.
    endpoint_aliases = {
        "index": ("ui.index", "/"),
        "inventario_form": ("ui.inventario_form", "/inventario-form"),
        "inventario_list": ("ui.inventario_list", "/inventario-list"),
        "ajustes": ("ui.ajustes", "/ajustes"),
        "informe": ("ui.informe", "/informe"),
    }
    for legacy_endpoint, (blueprint_endpoint, route_path) in endpoint_aliases.items():
        if legacy_endpoint in app.view_functions:
            continue
        app.add_url_rule(
            route_path,
            endpoint=legacy_endpoint,
            view_func=app.view_functions[blueprint_endpoint],
        )

    @app.errorhandler(404)
    def pagina_no_encontrada(_error):
        return render_template("404.html"), 404

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True, use_reloader=False)

