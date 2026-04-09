from datetime import datetime

from flask import Blueprint, render_template

from database.controller import get_dashboard_stats


ui_bp = Blueprint("ui", __name__)


@ui_bp.route("/")
def index():
    stats = get_dashboard_stats()
    data = {
        "year": datetime.now().year,
        **stats,
    }
    return render_template("index.html", data=data)


@ui_bp.route("/inventario-form")
def inventario_form():
    return inventario_list()


@ui_bp.route("/inventario-list")
def inventario_list():
    return render_template("inventario-list.html", data={"year": datetime.now().year})


@ui_bp.route("/ajustes")
def ajustes():
    return render_template("ajustes.html", data={"year": datetime.now().year})


@ui_bp.route("/informe")
def informe():
    return render_template("informe.html", data={"year": datetime.now().year})
