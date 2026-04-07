import shutil
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) in sys.path:
    sys.path.remove(str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR))

from flask import Flask, render_template

try:
    from app.blueprints import register_blueprints
except ModuleNotFoundError:
    from blueprints import register_blueprints

from database.controller import init_schema
from database.db import init_app


def _resolve_database_path(base_dir: Path) -> Path:
    target_db = base_dir / "inventario.db"
    legacy_db = base_dir / "prueba.db"

    if target_db.exists():
        return target_db

    if legacy_db.exists():
        try:
            legacy_db.replace(target_db)
        except OSError:
            shutil.copy2(legacy_db, target_db)

    return target_db


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["DATABASE"] = _resolve_database_path(BASE_DIR)
    init_app(app)

    with app.app_context():
        init_schema(BASE_DIR)

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
