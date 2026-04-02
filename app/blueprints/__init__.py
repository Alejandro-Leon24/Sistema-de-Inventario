from .ui import ui_bp
from .locations import locations_bp
from .inventory import inventory_bp
from .parameters import parameters_bp
from .administration import administration_bp
from .documents import documents_bp
from .files import files_bp
from .realtime import realtime_bp


ALL_BLUEPRINTS = [
    ui_bp,
    locations_bp,
    inventory_bp,
    parameters_bp,
    administration_bp,
    documents_bp,
    files_bp,
    realtime_bp,
]


def register_blueprints(app):
    for blueprint in ALL_BLUEPRINTS:
        app.register_blueprint(blueprint)
