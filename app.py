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
            b'g:li(Vr(a,"G1"))};if((!o.b||!o.c)',
            b'g:li(Vr(a,"G1")),virtualColumn:li(Vr(a,"I1")).includes("virtual")?"I":li(Vr(a,"H1")).includes("virtual")?"H":li(Vr(a,"G1")).includes("virtual")?"G":null};if((!o.b||!o.c)',
        ),
        (
            b'o.g.includes("virtual")||o.g.includes("aprob")&&o.g.includes("segunda")||t.push("G1 debe indicar virtualizaci\xc3\xb3n o aprobaci\xc3\xb3n en segunda instancia.")',
            b'o.g.includes("virtual")||o.g.includes("aprob")&&o.g.includes("segunda")||t.push("G1 debe indicar virtualizaci\xc3\xb3n o aprobaci\xc3\xb3n en segunda instancia."),o.virtualColumn||t.push("La hoja Control debe incluir Virtualizacion en I1, H1 o G1 (formato anterior).")',
        ),
        (
            b'h.virtualizadosReportados=iv(a,`G${m}`)',
            b'h.virtualizadosReportados=o.virtualColumn?iv(a,`${o.virtualColumn}${m}`):null',
        ),
        (
            b'const T=RB(k);T==="otro"&&r.push(`Fila ${m}: valor de virtualizaci\xc3\xb3n no reconocido (\xe2\x80\x9c${k}\xe2\x80\x9d).`)',
            b'const T=o.virtualColumn?RB(v):"sin_info";T==="otro"&&r.push(`Fila ${m}: valor de virtualizaci\xc3\xb3n no reconocido (\xe2\x80\x9c${v}\xe2\x80\x9d).`)',
        ),
        (
            b'"G \xc2\xb7 Virtualizaci\xc3\xb3n":k',
            b'[`${o.virtualColumn} \xc2\xb7 Virtualizaci\xc3\xb3n`]:v',
        ),
        (
            b'const w=a[`D${m}`],E=Vr(a,`E${m}`),S=Vr(a,`F${m}`),k=Vr(a,`G${m}`);',
            b'const w=a[`D${m}`],E=Vr(a,`E${m}`),S=Vr(a,`F${m}`),k=Vr(a,`G${m}`),v=o.virtualColumn?Vr(a,`${o.virtualColumn}${m}`):null;',
        ),
        (
            b'if(!y&&!g&&!A&&!E&&!S&&!k&&!(w&&w.v!=null))continue',
            b'if(!y&&!g&&!A&&!E&&!S&&!k&&!v&&!(w&&w.v!=null))continue',
        ),
        (
            b'virtRaw:k',
            b'virtRaw:v',
        ),
        (
            b'return{iso:null,raw:lh(e.v),problem:"fecha escrita como texto (no se interpreta)"}}const Vr=',
            b'const t=lh(e.v),r=t.match(/^(\\d{1,2})[\\/-](\\d{1,2})[\\/-](\\d{4})$/);if(!r)return{iso:null,raw:t,problem:"fecha escrita como texto con formato no reconocido"};const n=Number(r[1]),a=Number(r[2]),i=Number(r[3]),o=new Date(Date.UTC(i,a-1,n));return o.getUTCFullYear()===i&&o.getUTCMonth()===a-1&&o.getUTCDate()===n?{iso:`${i}-${Zf(a)}-${Zf(n)}`,raw:t,text:!0}:{iso:null,raw:t,problem:"fecha escrita como texto inv\xc3\xa1lida"}}const Vr=',
        ),
        (
            b'fechaISO:_.iso,fechaRaw:_.raw,revision:D',
            b'fechaISO:_.iso,fechaRaw:_.raw,fechaTexto:_.text===!0,revision:D',
        ),
        (
            b'return{...r,fechaAlmacenada:n,fechaISO:Number(o)<=12?`${a}-${o}-${i}`:n}})}function BB',
            b'return{...r,fechaAlmacenada:n,fechaISO:r.fechaTexto?n:Number(o)<=12?`${a}-${o}-${i}`:n}})}function BB',
        ),
        (
            b'd+=f.get(y)??0,s(y)?h.push({fecha:y,dia:f.get(y)??0,acumulado:d}):(v+=f.get(y)??0,h.length&&(h[h.length-1].acumulado=d))',
            b'd+=f.get(y)??0,s(y)?h.push({fecha:y,dia:f.get(y)??0,acumulado:d}):(v+=f.get(y)??0,f.has(y)&&h.push({fecha:y,dia:f.get(y)??0,acumulado:d}))',
        ),
        (
            b'Barras: PAD por fecha de carga \xc2\xb7 l\xc3\xadnea: acumulado desde el inicio \xc2\xb7 solo d\xc3\xadas h\xc3\xa1biles de lunes a viernes; no se descuentan festivos.',
            b'Barras: PAD por fecha de carga \xc2\xb7 l\xc3\xadnea: acumulado desde el inicio \xc2\xb7 incluye fines de semana con cargas; el ritmo se mide en d\xc3\xadas h\xc3\xa1biles sin descontar festivos.',
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
            b'APROBADO/COMPLETO: ${h?.aprobadosReportados??a.aprobados} PAD. EN REVISION: 8 PAD.',
        ),
        (
            b'function WB(e,t,r,n){const a=Di(e,t);',
            b'function WB(e,t,r,n,h){const a=Di(e,t);',
        ),
        (
            b'Revisi\xc3\xb3n por experto: ${a.completas} completa${a.completas===1?"":"s"} (${br(a.completas,a.n)} de ${a.n}), ${a.parciales} parcial${a.parciales===1?"":"es"} y ${a.sinRevision} sin revisi\xc3\xb3n registrada.',
            b'REVISION EXPERTO: ${a.completas} PAD (${br(a.completas,a.n)} de ${a.n}). REVISIONES CON COMENTARIO: ${h?.cambiosReportados??a.cambios} PAD.',
        ),
        (
            b'i.push(a.virtConf?`Virtualizaci\xc3\xb3n confirmada en ${a.virtConf} PAD; ${a.virtSinInfo} sin informaci\xc3\xb3n registrada.`:`No hay virtualizaciones confirmadas en el archivo; ${a.virtSinInfo} de ${a.n} PAD no tienen informaci\xc3\xb3n de virtualizaci\xc3\xb3n registrada.`);',
            b'i.push(`Virtualizaci\xc3\xb3n: ${h?.virtualizadosReportados??a.virtConf} subidos; ${a.virtSinInfo} sin informaci\xc3\xb3n registrada.`);',
        ),
        (
            b'WB(T,D,a.goal,G)',
            b'WB(T,D,a.goal,G,e.headerTotals)',
        ),
        (
            b'label:"APROBADO/COMPLETO",value:B.aprobados,tone:"ok",def:`${br(B.aprobados,B.n)} de ${B.n} \xc2\xb7 observaci\xc3\xb3n \xe2\x80\x9cAprobado\xe2\x80\x9d`',
            b'label:"APROBADO/COMPLETO",value:e.headerTotals?.aprobadosReportados??B.aprobados,tone:"ok",def:`${br(e.headerTotals?.aprobadosReportados??B.aprobados,B.n)} de ${B.n} \xc2\xb7 Cantidad de PADs aprobados en segunda instancia`',
        ),
        (
            b'b.jsx("small",{children:"Observaci\xc3\xb3n \xe2\x80\x9cAprobado\xe2\x80\x9d. Revisar no equivale a aprobar."})',
            b'b.jsx("small",{children:"Cantidad de PADs aprobados en segunda instancia."})',
        ),
        (
            b'warnings:r,records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null},dateFormat:null}',
            b'warnings:r,records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,aprobadosReportados:null,cambiosReportados:null},dateFormat:null}',
        ),
        (
            b'i={records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null},dateFormat:null}',
            b'i={records:[],engineers:[],headerTotals:{padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,aprobadosReportados:null,cambiosReportados:null},dateFormat:null}',
        ),
        (
            b'const h={padsReportados:null,revisadosReportados:null,virtualizadosReportados:null};',
            b'const h={padsReportados:null,revisadosReportados:null,virtualizadosReportados:null,aprobadosReportados:null,cambiosReportados:null};',
        ),
        (
            b'h.padsReportados=iv(a,`B${m}`),h.revisadosReportados=iv(a,`E${m}`),h.virtualizadosReportados=o.virtualColumn?iv(a,`${o.virtualColumn}${m}`):null;continue',
            b'h.padsReportados=iv(a,`B${m}`),h.revisadosReportados=iv(a,`E${m}`),h.virtualizadosReportados=o.virtualColumn?iv(a,`${o.virtualColumn}${m}`):null,h.aprobadosReportados=iv(a,`G${m}`),h.cambiosReportados=iv(a,`F${m}`);continue',
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
            b'Cantidad de PADs realizados por ingeniero \xc2\xb7 clic para filtrar',
        ),
        (
            b'b.jsx("div",{className:"legend",children:["completa","parcial","sin_registro"].map(ne=>b.jsxs("span",{children:[b.jsx("i",{style:{background:yp[ne]}}),va[ne]]},ne))})',
            b'b.jsx("div",{className:"legend",children:b.jsxs("span",{children:[b.jsx("i",{style:{background:je.green}}),"PAD realizados"]})})',
        ),
        (
            b'b.jsx("div",{className:"legend big",children:["completa","parcial","sin_registro"].map(x=>b.jsxs("span",{children:[b.jsx("i",{style:{background:yp[x]}}),va[x]]},x))})',
            b'b.jsx("div",{className:"legend big",children:b.jsxs("span",{children:[b.jsx("i",{style:{background:je.green}}),"PAD realizados"]})})',
        ),
        (
            b'b.jsx("b",{children:p.completas}),b.jsx("span",{children:"revisiones completas"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.parciales}),b.jsx("span",{children:"revisi\xc3\xb3n parcial"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.aprobados}),b.jsx("span",{children:"aprobaciones expl\xc3\xadcitas"})]}),b.jsxs("div",{children:[b.jsx("b",{children:p.observaciones}),b.jsx("span",{children:"con observaciones"})]})',
            b'b.jsx("b",{children:p.completas}),b.jsx("span",{children:"REVISION EXPERTO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:a.headerTotals?.cambiosReportados??p.cambios}),b.jsx("span",{children:"REVISIONES CON COMENTARIO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:a.headerTotals?.aprobadosReportados??p.aprobados}),b.jsx("span",{children:"APROBADO/COMPLETO"})]}),b.jsxs("div",{children:[b.jsx("b",{children:8}),b.jsx("span",{children:"EN REVISION"})]})',
        ),
        (
            b'["Revisiones completas",o.completas],["Parciales",o.parciales],["Sin revisi\xc3\xb3n",o.sinRevision],["Aprobados",o.aprobados],["Con observaciones",o.observaciones],["Virtualizaci\xc3\xb3n confirmada",o.virtConf]',
            b'["REVISION EXPERTO",o.completas],["REVISIONES CON COMENTARIO",r.headerTotals?.cambiosReportados??o.cambios],["Sin revisi\xc3\xb3n registrada",o.sinRevision,"Celda vac\xc3\xada."],["APROBADO/COMPLETO",r.headerTotals?.aprobadosReportados??o.aprobados],["EN REVISION",8],["VIRTUALIZACI\xc3\x93N \xc2\xb7 subidos",r.headerTotals?.virtualizadosReportados??o.virtConf,"Cantidad de PADs subidos para virtualizaci\xc3\xb3n."],["VIRTUALIZACI\xc3\x93N \xc2\xb7 sin informaci\xc3\xb3n",o.virtSinInfo,"PAD sin informaci\xc3\xb3n registrada."]',
        ),
        (
            b'].map(([c,u])=>b.jsxs("td",{children:[b.jsx("b",{children:u}),b.jsx("span",{children:c})]},String(c)))',
            b'].map(([c,u,d])=>b.jsxs("td",{children:[b.jsx("b",{children:u}),b.jsx("span",{children:c}),d&&b.jsx("small",{style:{display:"block",fontSize:"7pt",color:"#7A8A81"},children:d})]},String(c)))',
        ),
        (
            b',["Virtualizaci\xc3\xb3n en proceso",2]',
            b'',
        ),
        (
            b'b.jsx("h2",{children:"PADs en proceso de virtualizaci\\xf3n"}),b.jsxs("ul",{children:[b.jsx("li",{children:"Agroecoturismo"}),b.jsx("li",{children:"Redes Sociales en Marketing Digital"})]})',
            b'',
        ),
        (
            b'children:p.aprobados}),b.jsx("span",{children:"Aprobaci\xc3\xb3n expl\xc3\xadcita"}),b.jsx("small",{children:"Cantidad de PADs aprobados en segunda instancia."})]})',
            b'children:a.headerTotals?.aprobadosReportados??p.aprobados}),b.jsx("span",{children:"APROBADO/COMPLETO"}),b.jsx("small",{children:"Cantidad de PADs aprobados en segunda instancia."})]}),b.jsxs("div",{children:[b.jsx("b",{children:2}),b.jsx("span",{children:"Virtualizaci\xc3\xb3n"}),b.jsxs("small",{children:["en proceso de virtualizaci\xc3\xb3n.",b.jsx("br",{}),"Agroecoturismo"," \xc2\xb7 ","Redes Sociales en Marketing Digital"]})]})',
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
            b'b.jsx("span",{className:"k-def",children:r})',
            b'b.jsxs("span",{className:"k-def",children:[r&&r.includes("%")?r.split(" de ")[0]:null,r&&r.includes(" \xc2\xb7 ")?b.jsx("small",{style:{marginLeft:"4px"},children:r.split(" \xc2\xb7 ").slice(1).join(" \xc2\xb7 ")}):r&&!r.includes("%")?r:null]})',
        ),
        (
            b'b.jsxs("span",{className:"k-def",children:[br(ne.value,B.n)," de ",B.n]})',
            b'b.jsxs("span",{className:"k-def",children:[br(ne.value,B.n),b.jsx("small",{style:{marginLeft:"4px"},children:"PAD asignados a este ingeniero"})]})',
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
            b'value:8,tone:"alert",def:',
        ),
        (
            b'texto distinto de \xe2\x80\x9cAprobado\xe2\x80\x9d',
            b'Cantidad diaria de PADs en revisi\xc3\xb3n',
        ),
        (
            b'PAD asignados a este ingeniero',
            b'Cantidad de PADs asignados a este ingeniero',
        ),
    )
    for old, new in replacements:
        if html.count(old) != 1:
            raise RuntimeError(
                "No se encontró exactamente una referencia del tablero "
                f"que debía actualizarse: {old.decode('utf-8', errors='replace')}"
            )
        html = html.replace(old, new, 1)

    engineer_chart = (
        b'function M4({rs:e,engineers:t,f:r,setF:n,big:a}){'
        b'const o=[...t,...e.some(l=>l.ambos)?["Ambos (verificar)"]:[]]'
        b'.map(l=>({name:l,total:e.filter(c=>c.ingeniero===l).length})),'
        b's=a?260:170;return b.jsx(vf,{width:"100%",height:s,children:'
        b'b.jsxs(zy,{data:o,layout:"vertical",'
        b'margin:{left:8,right:40,top:4,bottom:4},barCategoryGap:a?28:18,'
        b'onClick:l=>(l==null?void 0:l.activeLabel)&&'
        b'n({ingeniero:r.ingeniero===l.activeLabel?"":l.activeLabel}),children:['
        b'b.jsx(Sf,{horizontal:!1,stroke:je.grid}),'
        b'b.jsx(Ki,{type:"number",allowDecimals:!1,tick:dl,axisLine:!1,tickLine:!1}),'
        b'b.jsx(Xi,{type:"category",dataKey:"name",width:a?150:104,'
        b'tick:{...dl,fontSize:a?16:13,fill:je.ink,fontWeight:600},'
        b'axisLine:!1,tickLine:!1}),'
        b'b.jsx(Ir,{...xp,formatter:(l)=>[l,"PAD realizados"]}),'
        b'b.jsx(qn,{dataKey:"total",name:"PAD realizados",fill:je.green,'
        b'cursor:"pointer",radius:[0,6,6,0],children:o.map(f=>'
        b'b.jsx(Hi,{opacity:r.ingeniero&&r.ingeniero!==f.name?.3:1},f.name))})'
        b']})})}'
    )
    html, chart_matches = re.subn(
        rb'function M4\(\{rs:e,engineers:t,f:r,setF:n,big:a\}\)\{.*?\}(?=function L4\()',
        engineer_chart,
        html,
        count=1,
        flags=re.DOTALL,
    )
    if chart_matches != 1:
        raise RuntimeError(
            "No se encontró exactamente una gráfica de producción por ingeniero."
        )

    removals = (
        (
            rb',h===4&&b\.jsxs\("div",\{className:"p-obs",children:\[.*?\}\),h===5&&',
            b',h===4&&',
            "la diapositiva de observaciones",
        ),
        (
            rb',\"Observaciones que requieren seguimiento\"',
            b"",
            "el título de la diapositiva de observaciones",
        ),
        (
            rb'b\.jsx\("h2",\{children:"Observaciones que requieren seguimiento"\}\),s\.length\?b\.jsxs\("table",\{children:\[.*?\}\):b\.jsx\("p",\{children:"Sin registros\."\}\),(?=b\.jsx\("h2",\{children:"Detalle por facultad y programa"\}\))',
            b"",
            "la sección de observaciones del informe imprimible",
        ),
    )
    for pattern, replacement, description in removals:
        html, matches = re.subn(pattern, replacement, html, count=1, flags=re.DOTALL)
        if matches != 1:
            raise RuntimeError(
                f"No se encontró exactamente una referencia para quitar {description}."
            )

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