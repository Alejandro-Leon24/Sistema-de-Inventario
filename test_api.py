import pytest

from app.app import app as flask_app, BASE_DIR
from database.schema_manager import init_schema


@pytest.fixture()
def client(tmp_path):
    db_path = tmp_path / "test_inventario.db"
    flask_app.config["TESTING"] = True
    flask_app.config["DATABASE"] = db_path

    with flask_app.app_context():
        init_schema(BASE_DIR)

    with flask_app.test_client() as test_client:
        yield test_client


def test_create_and_list_estados(client):
    payload = {"nombre": "Estado QA", "descripcion": "Creado por pytest"}
    create_response = client.post("/api/parametros/estados", json=payload)
    assert create_response.status_code in (201, 409)

    list_response = client.get("/api/parametros/estados")
    assert list_response.status_code == 200
    data = list_response.get_json()
    assert "data" in data
    assert isinstance(data["data"], list)
    assert any((row.get("nombre") or "").lower() == "estado qa" for row in data["data"])


def test_create_and_list_condiciones(client):
    payload = {"nombre": "Condicion QA", "descripcion": "Creado por pytest"}
    create_response = client.post("/api/parametros/condiciones", json=payload)
    assert create_response.status_code in (201, 409)

    list_response = client.get("/api/parametros/condiciones")
    assert list_response.status_code == 200
    data = list_response.get_json()
    assert "data" in data
    assert isinstance(data["data"], list)
    assert any((row.get("nombre") or "").lower() == "condicion qa" for row in data["data"])
