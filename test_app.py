import io
from pathlib import Path

import pytest

import app as appmod

BAJADA = Path(__file__).resolve().parent / "bajada_0610.xlsx"


@pytest.fixture
def cliente(monkeypatch, tmp_path):
    monkeypatch.setattr(appmod, "TMP", tmp_path / "cuarteo")
    monkeypatch.delenv("APP_PASSWORD", raising=False)
    appmod.app.config["TESTING"] = True
    return appmod.app.test_client()


def subir(cliente, datos=None, nombre="bajada.xlsx", **extra):
    contenido = datos if datos is not None else BAJADA.read_bytes()
    return cliente.post("/generar", data={"bajada": (io.BytesIO(contenido), nombre), **extra},
                        content_type="multipart/form-data")


def test_pagina_inicial(cliente):
    r = cliente.get("/")
    assert r.status_code == 200 and "Cuarteo final" in r.get_data(as_text=True)


def test_generar_y_descargar(cliente):
    r = subir(cliente, serie="2291")
    html = r.get_data(as_text=True)
    assert r.status_code == 200
    assert "Final generado" in html and "CUARTEO_0610.xlsx" in html
    assert "1.718" in html and "60,867" in html           # formato es-AR
    enlace = html.split('href="/descargar/')[1].split('"')[0]
    d = cliente.get(f"/descargar/{enlace}")
    assert d.status_code == 200
    assert d.data[:2] == b"PK"                             # es un xlsx (zip)
    assert "CUARTEO_0610.xlsx" in d.headers["Content-Disposition"]


def test_sin_archivo(cliente):
    r = cliente.post("/generar", data={}, content_type="multipart/form-data")
    assert r.status_code == 400 and "Elegí el archivo" in r.get_data(as_text=True)


def test_extension_incorrecta(cliente):
    r = subir(cliente, b"hola", nombre="bajada.pdf")
    assert r.status_code == 400 and ".xlsx" in r.get_data(as_text=True)


def test_archivo_invalido(cliente):
    r = subir(cliente, b"no es un excel")
    assert r.status_code == 400 and "No se pudo abrir" in r.get_data(as_text=True)


def test_serie_invalida(cliente):
    r = subir(cliente, serie="abc")
    assert r.status_code == 400 and "serie" in r.get_data(as_text=True)


def test_token_invalido(cliente):
    assert cliente.get("/descargar/../../etc/passwd").status_code == 404
    assert cliente.get("/descargar/corto").status_code == 404


def test_acceso_con_clave(cliente, monkeypatch):
    monkeypatch.setenv("APP_USER", "usuario")
    monkeypatch.setenv("APP_PASSWORD", "clave-de-prueba")
    assert cliente.get("/").status_code == 401
    assert cliente.get("/salud").status_code == 200
    from base64 import b64encode
    ok = {"Authorization": "Basic " + b64encode(b"usuario:clave-de-prueba").decode()}
    mal = {"Authorization": "Basic " + b64encode(b"usuario:otra").decode()}
    assert cliente.get("/", headers=ok).status_code == 200
    assert cliente.get("/", headers=mal).status_code == 401
