import base64
import json
import os
import re
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen

from flask import Flask, Response, jsonify, request


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
WORKBOOK_PATH = DATA_DIR / "seguimiento_pad.xlsx"
HTML_PATH = BASE_DIR / "templates" / "index.html"
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
SUPABASE_BUCKET = os.environ.get("SUPABASE_BUCKET", "seguimiento-pad")
SUPABASE_ENABLED = bool(SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY)
if "RENDER" in os.environ and not SUPABASE_ENABLED:
    raise RuntimeError(
        "En Render configura SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY."
    )
if (
    bool(SUPABASE_URL) != bool(SUPABASE_SERVICE_ROLE_KEY)
):
    raise RuntimeError(
        "Configura SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY."
    )

WORKBOOK_DATA_URI = re.compile(
    rb'const l0e="data:application/octet-stream;base64,[^"]+"'
)
INITIAL_LOAD = b"e||fetch(l0e)"
SUPABASE_OBJECT_URL = (
    f"{SUPABASE_URL}/storage/v1/object/"
    f"{quote(SUPABASE_BUCKET, safe='')}/seguimiento_pad.xlsx"
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


def supabase_headers(content_type: str | None = None) -> dict[str, str]:
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
    }
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def supabase_workbook() -> bytes | None:
    storage_request = Request(
        SUPABASE_OBJECT_URL, headers=supabase_headers()
    )
    try:
        with urlopen(storage_request, timeout=30) as response:
            return response.read()
    except HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        try:
            error_data = json.loads(error_body)
        except json.JSONDecodeError:
            error_data = {}
        error_message = str(error_data.get("message", error_body))
        normalized_message = error_message.lower()
        if (
            exc.code in (400, 404)
            and "bucket" not in normalized_message
            and (
                "object not found" in normalized_message
                or "resource not found" in normalized_message
            )
        ):
            return None
        raise RuntimeError(
            f"Supabase Storage devolvió HTTP {exc.code} al leer el Excel: "
            f"{error_message or 'sin detalle'}"
        ) from exc


def save_to_supabase(workbook: bytes, content_type: str) -> None:
    storage_request = Request(
        SUPABASE_OBJECT_URL,
        data=workbook,
        headers=supabase_headers(content_type)
        | {"x-upsert": "true"},
        method="POST",
    )
    try:
        with urlopen(storage_request, timeout=60):
            pass
    except HTTPError as exc:
        raise RuntimeError(
            f"Supabase Storage devolvió HTTP {exc.code} al guardar el Excel."
        ) from exc


def dashboard_html() -> bytes:
    html = HTML_PATH.read_bytes()
    if SUPABASE_ENABLED:
        workbook = supabase_workbook()
    elif WORKBOOK_PATH.is_file():
        workbook = WORKBOOK_PATH.read_bytes()
    else:
        workbook = None
    if workbook is None:
        return html

    workbook_data = base64.b64encode(workbook)
    replacement = (
        b'const l0e="data:application/octet-stream;base64,'
        + workbook_data
        + b'"'
    )
    html, workbook_matches = WORKBOOK_DATA_URI.subn(replacement, html)
    if workbook_matches != 1:
        raise RuntimeError(
            "No se encontró exactamente una fuente de Excel en el HTML del tablero."
        )

    html, load_matches = re.subn(
        re.escape(INITIAL_LOAD), b"fetch(l0e)", html, count=1
    )
    if load_matches != 1:
        raise RuntimeError(
            "No se encontró el punto de carga inicial de datos en el HTML del tablero."
        )
    return html


@app.get("/")
def index() -> Response:
    resp = Response(dashboard_html(), content_type="text/html; charset=utf-8")
    resp.headers["Cache-Control"] = "no-store"  # que todos vean siempre la última versión
    return resp


@app.post("/upload")
def upload():
    archivo = request.files.get("file")
    if archivo is None or not archivo.filename.lower().endswith(
        (".xlsx", ".xlsm", ".xls")
    ):
        return jsonify(error="Se requiere un archivo Excel .xlsx, .xlsm o .xls"), 400

    workbook = archivo.read()
    if SUPABASE_ENABLED:
        content_types = {
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
            ".xls": "application/vnd.ms-excel",
        }
        try:
            save_to_supabase(
                workbook, content_types[Path(archivo.filename).suffix.lower()]
            )
        except RuntimeError as exc:
            return jsonify(error=str(exc)), 502
    else:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        tmp = WORKBOOK_PATH.with_suffix(".tmp")
        tmp.write_bytes(workbook)
        os.replace(tmp, WORKBOOK_PATH)
    return jsonify(ok=True)


@app.get("/health")
def health() -> Response:
    return Response("ok", content_type="text/plain; charset=utf-8")


@app.get("/api/status")
def storage_status():
    if not SUPABASE_ENABLED:
        return jsonify(storage="local", configured=False, workbook_exists=False)

    try:
        workbook = supabase_workbook()
    except RuntimeError as exc:
        return jsonify(storage="supabase", configured=True, error=str(exc)), 502

    return jsonify(
        storage="supabase",
        configured=True,
        bucket=SUPABASE_BUCKET,
        workbook_exists=workbook is not None,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)