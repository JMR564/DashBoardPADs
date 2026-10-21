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
        SUPABASE_OBJECT_URL,
        headers=supabase_headers() | {"Cache-Control": "no-cache, no-store"},
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
    replacements = (
        (
            b'o.g.includes("virtual")||t.push("G1 debe ser la columna de virtualizaci\xc3\xb3n.")',
            b'o.g.includes("virtual")||o.g.includes("aprob")&&o.g.includes("segunda")||t.push("G1 debe indicar virtualizaci\xc3\xb3n o aprobaci\xc3\xb3n en segunda instancia.")',
        ),
        (
            b'h.virtualizadosReportados=iv(a,`G${m}`)',
            b'h.virtualizadosReportados=o.g.includes("virtual")?iv(a,`G${m}`):null',
        ),
        (
            b'const T=RB(k);T==="otro"&&r.push(`Fila ${m}: valor de virtualizaci\xc3\xb3n no reconocido (\xe2\x80\x9c${k}\xe2\x80\x9d).`)',
            b'const T=o.g.includes("virtual")?RB(k):"sin_info";T==="otro"&&r.push(`Fila ${m}: valor de virtualizaci\xc3\xb3n no reconocido (\xe2\x80\x9c${k}\xe2\x80\x9d).`)',
        ),
        (
            b'"G \xc2\xb7 Virtualizaci\xc3\xb3n":k',
            b'[o.g.includes("virtual")?"G \xc2\xb7 Virtualizaci\xc3\xb3n":"G \xc2\xb7 Aprobado en segunda instancia"]:k',
        ),
        (
            b'udec-pad:v1:data',
            b'udec-pad:v2:data',
        ),
        (
            b'DA={aprobado:"Aprobado (expl\xc3\xadcito)",observacion:"Con observaciones",sin_observacion:"Sin observaci\xc3\xb3n registrada"}',
            b'DA={aprobado:"Aprobado en segunda instancia",observacion:"En revisi\xc3\xb3n (segunda instancia)",sin_observacion:"Sin estado en segunda instancia"}',
        ),
        (
            b'Aprobaci\xc3\xb3n expl\xc3\xadcita (\xe2\x80\x9cAprobado\xe2\x80\x9d) en ${a.aprobados} PAD. ${a.observaciones} PAD tienen observaciones distintas de \xe2\x80\x9cAprobado\xe2\x80\x9d; el archivo no registra si ya fueron atendidas.',
            b'Aprobaci\xc3\xb3n en segunda instancia (\xe2\x80\x9cAprobado\xe2\x80\x9d) en ${a.aprobados} PAD. ${a.observaciones} PAD tienen otro valor registrado en esa columna.',
        ),
        (
            b'label:"APROBADO/COMPLETO",value:B.aprobados,tone:"ok",def:`${br(B.aprobados,B.n)} de ${B.n} \xc2\xb7 observaci\xc3\xb3n \xe2\x80\x9cAprobado\xe2\x80\x9d`',
            b'label:"APROBADO/COMPLETO",value:B.aprobados,tone:"ok",def:`${br(B.aprobados,B.n)} de ${B.n} \xc2\xb7 Cantidad de PADs aprobados en segunda instancia`',
        ),
        (
            b'b.jsx("small",{children:"Observaci\xc3\xb3n \xe2\x80\x9cAprobado\xe2\x80\x9d. Revisar no equivale a aprobar."})',
            b'b.jsx("small",{children:"Cantidad de PADs aprobados en segunda instancia."})',
        ),
        (
            b'warnings:r,records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null},dateFormat:null}',
            b'warnings:r,records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,cambiosReportados:null},dateFormat:null}',
        ),
        (
            b'i={records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null},dateFormat:null}',
            b'i={records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,cambiosReportados:null},dateFormat:null}',
        ),
        (
            b'const h={padsReportados:null,revisadosReportados:null,virtualizadosReportados:null};',
            b'const h={padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,cambiosReportados:null};',
        ),
        (
            b'h.padsReportados=iv(a,`B${m}`),h.revisadosReportados=iv(a,`E${m}`),h.virtualizadosReportados=o.g.includes("virtual")?iv(a,`G${m}`):null;continue',
            b'h.padsReportados=iv(a,`B${m}`),h.revisadosReportados=iv(a,`E${m}`),h.virtualizadosReportados=o.g.includes("virtual")?iv(a,`G${m}`):null,h.cambiosReportados=iv(a,`F${m}`);continue',
        ),
        (
            b'label:"REVISIONES CON COMENTARIO",value:B.cambios,',
            b'label:"REVISIONES CON COMENTARIO",value:e.headerTotals?.cambiosReportados??B.cambios,',
        ),
        (
            b'aprobados:ea(e,n=>n.obs==="aprobado")',
            b'aprobados:ea(e,n=>["si","s","yes"].includes(li(n.raw["G \xc2\xb7 Aprobado en segunda instancia"]))||jB(n.raw["G \xc2\xb7 Aprobado en segunda instancia"])==="aprobado")',
        ),
        (
            b'(!t.obs||n.obs===t.obs)&&(!t.cambios',
            b'(!t.obs||n.obs===t.obs)&&(!t.aprobacion||jB(n.raw["G \xc2\xb7 Aprobado en segunda instancia"])===t.aprobacion)&&(!t.cambios',
        ),
        (
            b'obs:"",virt:""',
            b'obs:"",aprobacion:"",virt:""',
        ),
        (
            b'onClick:()=>X({obs:"aprobado"}),on:o.obs==="aprobado"',
            b'onClick:()=>X({aprobacion:"aprobado"}),on:o.aprobacion==="aprobado"',
        ),
        (
            b'function M4({rs:e',
            b'const revisionChartLabels={completa:"REVISION EXPERTO",parcial:"APROBADO/COMPLETO"};function M4({rs:e',
        ),
        (
            b'parcial:c.filter(u=>u.revision==="parcial").length',
            b'parcial:c.filter(u=>["si","s","yes"].includes(li(u.raw["G \xc2\xb7 Aprobado en segunda instancia"]))||jB(u.raw["G \xc2\xb7 Aprobado en segunda instancia"])==="aprobado").length',
        ),
        (
            b'formatter:(l,c)=>[l,va[c]??c]',
            b'formatter:(l,c)=>[l,revisionChartLabels[c]??va[c]??c]',
        ),
        (
            b'name:va[i],value:e.filter',
            b'name:revisionChartLabels[i]??va[i],value:e.filter',
        ),
        (
            b'dataKey:l,stackId:"a",fill:yp[l]',
            b'dataKey:l,fill:yp[l]',
        ),
        (
            b',l==="sin_registro"&&b.jsx(Hr,{dataKey:"total",position:"right",style:{fontWeight:700,fill:je.ink,fontSize:a?18:13}})]},l))',
            b']},l))',
        ),
        (
            b'Segmentos por estado de revisi\xc3\xb3n \xc2\xb7 clic para filtrar',
            b'Conteos por indicador \xc2\xb7 pueden coincidir \xc2\xb7 clic para filtrar',
        ),
        (
            b'b.jsx("b",{children:p.completas}),b.jsx("span",{children:"revisiones completas"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.parciales}),b.jsx("span",{children:"revisi\xc3\xb3n parcial"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.aprobados}),b.jsx("span",{children:"aprobaciones expl\xc3\xadcitas"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.observaciones}),b.jsx("span",{children:"con observaciones"})]})',
            b'b.jsx("b",{children:p.completas}),b.jsx("span",{children:"REVISION EXPERTO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.cambios}),b.jsx("span",{children:"REVISIONES CON COMENTARIO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.aprobados}),b.jsx("span",{children:"APROBADO/COMPLETO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:20}),b.jsx("span",{children:"EN REVISION"})]})',
        ),
        (
            b'b.jsx("b",{children:p.completas}),b.jsxs("span",{children:["Revisi\xc3\xb3n completa \xc2\xb7 ",br(p.completas,p.n)," de ",p.n]})',
            b'b.jsx("b",{children:p.completas}),b.jsx("span",{children:"REVISION EXPERTO"})',
        ),
        (
            b'b.jsx("b",{children:p.parciales}),b.jsx("span",{children:"Revisi\xc3\xb3n parcial"}),b.jsx("small",{children:"Valor \xe2\x80\x9cA medias\xe2\x80\x9d."})',
            b'b.jsx("b",{children:p.cambios}),b.jsx("span",{children:"REVISIONES CON COMENTARIO"}),b.jsx("small",{children:"Cantidad de PADs con comentarios."})',
        ),
        (
            b'va[ne]]},ne))',
            b'revisionChartLabels[ne]??va[ne]]},ne))',
        ),
        (
            b'va[x]]},x))',
            b'revisionChartLabels[x]??va[x]]},x))',
        ),
        (
            b'b.jsx("span",{className:"k-def",children:r})',
            b'b.jsxs("span",{className:"k-def",children:[r&&r.includes("%")?r.split(" de ")[0]:null,r&&r.includes(" \xc2\xb7 ")?b.jsx("small",{style:{marginLeft:"4px"},children:r.split(" \xc2\xb7 ").slice(1).join(" \xc2\xb7 ")}):r&&!r.includes("%")?r:null]})',
        ),
        (
            b'b.jsxs("span",{className:"k-def",children:[br(ne.value,B.n)," de ",B.n]})',
            b'b.jsxs("span",{className:"k-def",children:[br(ne.value,B.n),b.jsx("small",{style:{marginLeft:"4px"},children:"PAD asignados a este ingeniero"})]})',
        ),
        (
            b'b.jsxs("span",{className:"k-def",children:["Denominador: ",B.n," PAD. Celda vac\xc3\xada = sin informaci\xc3\xb3n, no \xe2\x80\x9cpendiente\xe2\x80\x9d."]})',
            b'b.jsx("span",{className:"k-def",children:"Estado de virtualizaci\xc3\xb3n en el archivo"})',
        ),
        (
            b'def:G?`de ${_.length} en el archivo`:`${B.programas} programas \xc2\xb7 ${B.facultades} facultades`',
            b'def:"Cantidad total de PADs registrados"',
        ),
        (
            b'valor \xe2\x80\x9cS\xc3\xad\xe2\x80\x9d en revisi\xc3\xb3n',
            b'Cantidad de PADs revisados por experto',
        ),
        (
            b'con texto en Cambios a realizar',
            b'Cantidad de PADs con comentarios por actualizar',
        ),
        (
            b'celda de revisi\xc3\xb3n vac\xc3\xada',
            b'Cantidad de PADs sin revisi\xc3\xb3n',
        ),
        (
            b'value:20,tone:"alert",def:',
            b'value:10,tone:"alert",def:',
        ),
        (
            b'texto distinto de \xe2\x80\x9cAprobado\xe2\x80\x9d',
            b'Cantidad diaria de PADs en revisi\xc3\xb3n',
        ),
        (
            b'PAD asignados a este ingeniero',
            b'Cantidad de PADs asignados a este ingeniero',
        ),
        (
            b'Estado de virtualizaci\xc3\xb3n en el archivo',
            b'Cantidad de PADs seg\xc3\xban su estado de virtualizaci\xc3\xb3n',
        ),
    )
    for old, new in replacements:
        if html.count(old) != 1:
            raise RuntimeError(
                "No se encontró exactamente una referencia del tablero "
                f"que debía actualizarse: {old.decode('utf-8', errors='replace')}"
            )
        html = html.replace(old, new, 1)

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