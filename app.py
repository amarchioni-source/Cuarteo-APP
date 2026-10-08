"""App web del CUARTEO final: se sube la bajada del sistema y se descarga el final."""
import datetime as dt
import hmac
import io
import logging
import os
import re
import secrets
import time
from pathlib import Path

from flask import (Flask, Response, abort, redirect, render_template, request,
                   send_file, url_for)

import generar_cuarteo as g

BASE = Path(__file__).resolve().parent
PLANTILLA = BASE / "CUARTEO_plantilla.xlsx"
TMP = Path(os.environ.get("TMPDIR", "/tmp")) / "cuarteo"
VIDA_ARCHIVOS = 2 * 3600            # los finales generados se borran a las 2 horas
TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{16,64}$")

app = Flask(__name__, template_folder=str(BASE), static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024      # 8 MB
log = logging.getLogger("cuarteo")


# ------------------------------------------------------------------- acceso
@app.before_request
def requerir_clave():
    """Si APP_PASSWORD está definida (variable de entorno), pide usuario y clave."""
    clave = os.environ.get("APP_PASSWORD")
    if not clave or request.path == "/salud":
        return None
    usuario = os.environ.get("APP_USER", "")
    a = request.authorization
    ok = (a is not None and a.type == "basic"
          and hmac.compare_digest((a.username or "").encode(), usuario.encode())
          and hmac.compare_digest((a.password or "").encode(), clave.encode()))
    if ok:
        return None
    return Response("Acceso restringido.", 401,
                    {"WWW-Authenticate": 'Basic realm="Cuarteo", charset="UTF-8"'})


@app.after_request
def cabeceras(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ---------------------------------------------------------------- utilidades
@app.template_filter("num")
def num(valor, decimales=0):
    """Formato es-AR: 1.234,56"""
    if valor is None:
        return "—"
    s = f"{valor:,.{decimales}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


@app.template_filter("fecha")
def fecha(valor):
    return valor.strftime("%d/%m/%Y") if valor else ""


def limpiar_viejos():
    TMP.mkdir(parents=True, exist_ok=True)
    limite = time.time() - VIDA_ARCHIVOS
    for p in TMP.glob("*.xlsx"):
        try:
            if p.stat().st_mtime < limite:
                p.unlink()
        except OSError:
            pass


def render_index(status=200, **ctx):
    return render_template("index.html", form=request.form, **ctx), status


# -------------------------------------------------------------------- rutas
@app.get("/salud")
def salud():
    return {"estado": "ok"}


@app.get("/")
def index():
    return render_index()


@app.get("/generar")
def generar_get():
    """Quien abre /generar directamente (enlace copiado, recarga) vuelve al inicio."""
    return redirect(url_for("index"))


@app.post("/generar")
def generar():
    archivo = request.files.get("bajada")
    if archivo is None or not archivo.filename:
        return render_index(400, error="Elegí el archivo de la bajada para continuar.")
    if not archivo.filename.lower().endswith(".xlsx"):
        return render_index(400, error="El archivo debe ser un Excel .xlsx.")

    ff = None
    txt = (request.form.get("fecha_faena") or "").strip()
    if txt:
        try:
            ff = dt.date.fromisoformat(txt)
        except ValueError:
            return render_index(400, error="La fecha de faena no es válida.")
    serie = None
    txt = (request.form.get("serie") or "").strip()
    if txt:
        if not txt.isdigit() or len(txt) > 8:
            return render_index(400, error="El número de serie debe ser numérico.")
        serie = int(txt)

    salida = io.BytesIO()
    try:
        res = g.generar(io.BytesIO(archivo.read()), PLANTILLA, salida, ff, serie)
    except g.BajadaInvalida as e:
        return render_index(400, error=str(e))
    except Exception:                                    # noqa: BLE001
        log.exception("Error inesperado al generar el cuarteo")
        return render_index(500, error="No se pudo generar el final. "
                                       "Revisá que el archivo sea la bajada original del sistema.")

    limpiar_viejos()
    token = secrets.token_urlsafe(18)
    nombre = f"CUARTEO_{res['fecha']:%d%m}.xlsx"
    (TMP / f"{token}.xlsx").write_bytes(salida.getvalue())
    (TMP / f"{token}.nombre").write_text(nombre, encoding="utf-8")
    return render_index(res=res, token=token, nombre=nombre)


@app.get("/descargar/<token>")
def descargar(token):
    if not TOKEN_RE.match(token):
        abort(404)
    ruta = TMP / f"{token}.xlsx"
    if not ruta.is_file():
        abort(404)
    try:
        nombre = (TMP / f"{token}.nombre").read_text(encoding="utf-8")
    except OSError:
        nombre = "CUARTEO.xlsx"
    return send_file(ruta, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.errorhandler(413)
def demasiado_grande(_e):
    return render_index(413, error="El archivo es demasiado grande (máximo 8 MB).")


@app.errorhandler(404)
def no_encontrado(_e):
    return render_template("index.html", form={}, error="El enlace de descarga venció. "
                           "Volvé a generar el final."), 404


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
