import base64
from pathlib import Path
import re

from flask import Flask, Response, render_template


BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = BASE_DIR / "data" / "seguimiento_pad.xlsx"
HTML_PATH = BASE_DIR / "templates" / "index.html"
WORKBOOK_DATA_URI = re.compile(
    rb'const l0e="data:application/octet-stream;base64,[^"]+"'
)
INITIAL_LOAD = b"e||fetch(l0e)"

app = Flask(__name__)


def dashboard_html() -> bytes:
    html = HTML_PATH.read_bytes()
    if not WORKBOOK_PATH.is_file():
        return html

    workbook_data = base64.b64encode(WORKBOOK_PATH.read_bytes())
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
        re.escape(INITIAL_LOAD),
        b"fetch(l0e)",
        html,
        count=1,
    )
    if load_matches != 1:
        raise RuntimeError(
            "No se encontró el punto de carga inicial de datos en el HTML del tablero."
        )
    return html


@app.get("/")
def index() -> Response:
    return Response(
        dashboard_html(),
        content_type="text/html; charset=utf-8",
    )


@app.get("/health")
def health() -> Response:
    return Response("ok", content_type="text/plain; charset=utf-8")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
