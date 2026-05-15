#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
conversor_papel.py v11

Convierte una unidad online DOCX/PDF/TXT a DOCX papel editorial usando como base
un DOCX ya maquetado y, opcionalmente, un DOCX de interacciones.

Uso:
  python conversor_papel.py UNIDAD.docx EJEMPLO_MAQUETADO.docx SALIDA.docx
  python conversor_papel.py UNIDAD.docx EJEMPLO_MAQUETADO.docx INTERACCIONES.docx SALIDA.docx
  python conversor_papel.py UNIDAD.docx EJEMPLO_MAQUETADO.docx PLANTILLA.docx INTERACCIONES.docx SALIDA.docx

Cambios clave v10:
  - Inserta gráficos reales del DOCX de referencia: imágenes, SmartArt, diagramas, dibujos.
  - Corrige el problema de namespace wp14: los SmartArt usaban wp14:anchorId y no se declaraba.
  - Copia TODO el paquete del ejemplo salvo document.xml, para que los rId y parts de SmartArt funcionen.
  - Extrae párrafos con regex correcta: no confunde <w:pPr> con <w:p>.
  - Usa los namespace reales del documento de referencia.
  - Conserva cursivas/negritas de los runs y muestra en papel la URL real de los hipervínculos.
  - Respeta la numeración visible del online en actividades y tareas.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape


# =============================================================================
# Regex y constantes
# =============================================================================

RE_URL = re.compile(r"^https?://", re.IGNORECASE)
RE_SEC1 = re.compile(r"^(\d+)\.\s+(.+)")
RE_SEC2 = re.compile(r"^(\d+)\.(\d+)\s+(.+)")
RE_SEC3 = re.compile(r"^(\d+)\.(\d+)\.(\d+)\s+(.+)")
RE_INTER = re.compile(r"^Interacci[oó]n\s+(\d+)(?:\.?\s+(.+))?$", re.IGNORECASE)
RE_OPCION = re.compile(r"^([a-h])\)\s+(.+)")
RE_FORMULA = re.compile(
    r"(?:"
    r"\d+\s*[×x\*/÷]\s*\d+"
    r"|[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+\s*=\s*.+[/÷×]"
    r"|\b\w+\s*/\s*\w+\s*[×x]\s*100"
    r")"
)

RE_PUA = re.compile(r"[\ue000-\uf8ff\U000f0000-\U000fffff]+")
RE_PUA_ONLY = re.compile(r"^[\ue000-\uf8ff\U000f0000-\U000fffff\s]+$")

BLOQUES_ESP = {
    "Nota", "Ejemplo", "Sabías que...", "Sabías que…", "Consejo",
    "Definición", "Hilo conductor", "Para saber más", "Vídeo",
    "Importante", "Recuerda",
}

PREFIJOS_ELIM = (
    "Para realizar las Actividades colaborativas",
    "Para realizar las Tareas de evaluación",
    "Las instrucciones para realizar la tarea",
    "Podrás compartir ",
    "Podrás debatir",
    "poder debatir y aportar",
    "Podrás identificar las",
    "Es el momento de realizar la siguiente",
    "No obstante puedes seguir estudiando",
    "La duración aproximada de la misma",
    "Mapa conceptual o esquema de contenidos",
    "Para realizar las",
    "valoración será tenida en cuenta",
    "encontrarás la información necesaria",
    "otro momento que te sea más favorable",
    "Cambio de pantalla",
    "Pulsa en ",
    "Pulsa para ",
    "Avanza para ",
    "Haz clic ",
    "Haz clic para ",
    "Haz clic en ",
    "Pincha ",
    "Instrucción:",
    "Instrucción: ",
)

RESIDUOS_RESUMEN = {
    "Conservación o eliminación",
    "Almacenamiento y custodia",
    "Uso y tramitación",
    "Clasificación o registro",
    "Creación o recepción",
    "Clasificación y registro",
}

VINETA_SIM = {1: "●", 2: "○", 3: "▪", 4: "–"}
VINETA_EST = {
    1: "Vietanvl11d",
    2: "Vietanvl21d",
    3: "Vietanvl31d",
    4: "Vietanvl41d",
}


# =============================================================================
# Helpers generales
# =============================================================================

def esc(texto: str) -> str:
    return xml_escape(str(texto), {'"': "&quot;"})


def limpiar_pua(texto: str) -> str:
    return RE_PUA.sub("", texto or "").strip()


def es_solo_pua(texto: str) -> bool:
    return bool(texto) and bool(RE_PUA_ONLY.match(texto))


def debe_elim(texto: str) -> bool:
    if not texto:
        return True
    if texto in RESIDUOS_RESUMEN:
        return True
    return any(texto.startswith(p) for p in PREFIJOS_ELIM)


def limpiar_titulo(texto: str) -> str:
    texto = re.sub(r"\s*\(CE\s+[a-z]\)\s*y\s*\(CE\s+[a-z]\)", "", texto, flags=re.I)
    texto = re.sub(r"\s*\(CE\s+[a-z…]+\)", "", texto, flags=re.I)
    texto = re.sub(r"\s*\(Ce[^)]*\)", "", texto, flags=re.I)
    texto = re.sub(r"\s*…+$", "", texto)
    texto = re.sub(r"\s*\.{2,}$", "", texto)
    return texto.rstrip("…. ").strip()


def infinitivo_a_imperativo(texto: str) -> str:
    pares = [
        ("buscar ", "Busca "), ("identificar ", "Identifica "),
        ("analizar ", "Analiza "), ("elaborar ", "Elabora "),
        ("diseñar ", "Diseña "), ("crear ", "Crea "),
        ("realizar ", "Realiza "), ("comparar ", "Compara "),
        ("describir ", "Describe "), ("explicar ", "Explica "),
        ("completar ", "Completa "), ("redactar ", "Redacta "),
        ("investigar ", "Investiga "), ("seleccionar ", "Selecciona "),
        ("clasificar ", "Clasifica "), ("calcular ", "Calcula "),
        ("consultar ", "Consulta "), ("revisar ", "Revisa "),
    ]
    low = texto.lower()
    for inf, imp in pares:
        if low.startswith(inf):
            return imp + texto[len(inf):]
    return texto[:1].upper() + texto[1:] if texto else texto


def _es_url_imagen(url: str) -> bool:
    dominios = (
        "shutterstock.com", "gettyimages.", "istockphoto.com", "unsplash.com",
        "pexels.com", "freepik.com", "pixabay.com", "depositphotos.com",
        "adobe.com/stock", "stock.adobe.com",
    )
    u = url.lower()
    if any(d in u for d in dominios):
        return True
    return bool(re.search(r"\.(jpg|jpeg|png|gif|webp|svg|bmp|tiff?)(\?|$)", u))


def _es_zip(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(4) == b"PK\x03\x04"
    except Exception:
        return False


def _style_name(p) -> str:
    try:
        return p.style.name
    except Exception:
        return ""


def _is_bold(p) -> bool:
    return any(r.bold for r in p.runs if r.text.strip())


def _norm_line(texto: str) -> str:
    texto = limpiar_pua(texto)
    texto = texto.replace("\x00", "")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def _limpiar_vineta_literal(texto: str) -> str:
    texto = texto.strip()
    texto = re.sub(r"^[●○▪\-–]\s*", "", texto)
    texto = re.sub(r"^[●○▪]\t", "", texto)
    return texto.strip()


def _abre_modo_lista(linea: str) -> bool:
    l = linea.strip().lower()
    if not l.endswith(":"):
        return False

    claves = (
        "permite", "mediante", "divide", "dividirse", "aspectos", "como",
        "son", "serían", "radica en", "se basa en", "incluye", "destacan",
        "siguientes", "pudiendo ser",
    )
    return any(c in l for c in claves)


def _parece_item_lista_en_bloque(linea: str, modo_lista: bool) -> bool:
    linea = linea.strip()

    if not linea:
        return False

    if re.match(r"^[●○▪\-–]\s+", linea):
        return True

    if not modo_lista:
        return False

    if len(linea) > 110:
        return False

    if linea.startswith((
        "Algunas ", "Alguna ", "En ", "Por ", "Para ", "Cuando ",
        "Una ", "Un ", "El ", "La ", "Los ", "Las ", "Este ", "Esta ",
        "Estos ", "Estas ", "Según ", "Asimismo ", "Además ",
    )):
        return False

    return True



# =============================================================================
# Texto enriquecido e hipervínculos
# =============================================================================

_HYPERLINKS_OUT: dict[str, str] = {}
_HYPERLINK_SEQ = 1


def _reset_hyperlinks_out() -> None:
    global _HYPERLINKS_OUT, _HYPERLINK_SEQ
    _HYPERLINKS_OUT = {}
    _HYPERLINK_SEQ = 1


def _rid_hyperlink(target: str) -> str:
    """Devuelve/crea un rId estable para un hipervínculo externo de salida."""
    global _HYPERLINK_SEQ
    target = str(target or "").strip()
    if not target:
        return ""
    if target not in _HYPERLINKS_OUT:
        _HYPERLINKS_OUT[target] = f"rIdHyper{_HYPERLINK_SEQ}"
        _HYPERLINK_SEQ += 1
    return _HYPERLINKS_OUT[target]


def rich_text(obj) -> str:
    if isinstance(obj, dict):
        return str(obj.get("texto", ""))
    return "" if obj is None else str(obj)


def rich_runs(obj):
    if isinstance(obj, dict):
        return obj.get("runs") or [{"text": rich_text(obj)}]
    return [{"text": rich_text(obj)}]


def rich_obj(texto: str, runs=None) -> dict:
    return {"texto": texto or "", "runs": runs or [{"text": texto or ""}]}


def _trim_runs(runs: list[dict]) -> list[dict]:
    runs = [dict(r) for r in (runs or []) if str(r.get("text", ""))]
    if not runs:
        return []
    runs[0]["text"] = str(runs[0].get("text", "")).lstrip()
    runs[-1]["text"] = str(runs[-1].get("text", "")).rstrip()
    return [r for r in runs if str(r.get("text", ""))]


def _materializar_hipervinculos_para_papel(runs: list[dict]) -> list[dict]:
    """
    En papel no debe quedar el texto ancla del online (p. ej. el título de un vídeo),
    sino la URL real. Conserva el hipervínculo en el DOCX generado y muestra la URL
    como texto visible.
    """
    out: list[dict] = []
    last_link = None

    for r in runs or []:
        rr = dict(r)
        link = str(rr.get("link", "")).strip()

        if link and RE_URL.match(link):
            # Si varios runs consecutivos forman el mismo hipervínculo, se sustituyen
            # por una sola URL visible para evitar duplicados.
            if link == last_link:
                continue
            rr["text"] = link
            rr["link"] = link
            # El estilo de negrita/cursiva del texto ancla no debe contaminar la URL.
            rr.pop("bold", None)
            rr.pop("italic", None)
            out.append(rr)
            last_link = link
        else:
            out.append(rr)
            last_link = None

    return _trim_runs(out)


def _runs_to_xml(runs: list[dict], force_bold: bool = False) -> str:
    partes: list[str] = []
    for r in _trim_runs(runs):
        txt = str(r.get("text", ""))
        if not txt:
            continue
        te = esc(txt)
        sp = ' xml:space="preserve"' if te and te != te.strip() else ""
        props = []
        if r.get("link"):
            props.append('<w:rStyle w:val="Hyperlink"/>')
        if force_bold or r.get("bold"):
            props.append("<w:b/><w:bCs/>")
        if r.get("italic"):
            props.append("<w:i/><w:iCs/>")
        rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
        run_xml = f"<w:r>{rpr}<w:t{sp}>{te}</w:t></w:r>"
        if r.get("link"):
            rid = _rid_hyperlink(str(r.get("link")))
            if rid:
                run_xml = f'<w:hyperlink r:id="{rid}" w:history="1">{run_xml}</w:hyperlink>'
        partes.append(run_xml)
    return "".join(partes)


def p_rich(obj, estilo: str, negrita: bool = False) -> str:
    texto = rich_text(obj)
    runs = rich_runs(obj)
    if not texto and not runs:
        return f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr></w:p>'
    return (
        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
        f"{_runs_to_xml(runs, force_bold=negrita)}</w:p>"
    )


def _map_runs_text(runs: list[dict], func) -> dict:
    texto = func("".join(str(r.get("text", "")) for r in runs))
    return rich_obj(texto, [{"text": texto}])


def _strip_literal_from_rich(obj, literal: str) -> dict:
    texto = rich_text(obj)
    runs = rich_runs(obj)
    if not texto.startswith(literal):
        return rich_obj(texto, runs)
    rest = texto[len(literal):].lstrip()
    return rich_obj(rest, [{"text": rest}])


def _prefix_rich(prefix: str, obj) -> dict:
    runs = [{"text": prefix}] + rich_runs(obj)
    return rich_obj(prefix + rich_text(obj), runs)


def _limpiar_vineta_rich(obj) -> dict:
    texto = rich_text(obj)
    limpio = _limpiar_vineta_literal(texto)
    if limpio == texto:
        return rich_obj(texto, rich_runs(obj))
    return rich_obj(limpio, [{"text": limpio}])


def _patch_document_rels_hyperlinks(rels_xml: bytes | str) -> bytes:
    xml = rels_xml.decode("utf-8") if isinstance(rels_xml, bytes) else str(rels_xml)
    if not _HYPERLINKS_OUT:
        return xml.encode("utf-8")
    inserts = []
    for target, rid in _HYPERLINKS_OUT.items():
        if f'Id="{rid}"' in xml:
            continue
        inserts.append(
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{esc(target)}" TargetMode="External"/>'
        )
    if inserts:
        xml = xml.replace("</Relationships>", "".join(inserts) + "</Relationships>")
    return xml.encode("utf-8")

# =============================================================================
# Constructores XML
# =============================================================================

def p(texto: str, estilo: str, negrita: bool = False) -> str:
    if isinstance(texto, dict):
        return p_rich(texto, estilo, negrita)

    texto = "" if texto is None else str(texto)
    te = esc(texto)
    sp = ' xml:space="preserve"' if te and te != te.strip() else ""

    if not te:
        return f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr></w:p>'

    rpr = "<w:rPr><w:b/><w:bCs/></w:rPr>" if negrita else ""

    return (
        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
        f"<w:r>{rpr}<w:t{sp}>{te}</w:t></w:r></w:p>"
    )


def p_vineta(texto: str, nivel: int = 1) -> str:
    estilo = VINETA_EST.get(nivel, VINETA_EST[1])
    simbolo = esc(VINETA_SIM.get(nivel, VINETA_SIM[1]))

    if isinstance(texto, dict):
        clean = _limpiar_vineta_literal(rich_text(texto))
        runs = rich_runs(texto)
        # Si había una viñeta literal al principio, la quitamos del texto enriquecido.
        if clean != rich_text(texto):
            runs = [{"text": clean}]
        return (
            f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
            f"<w:r><w:t>{simbolo}</w:t></w:r>"
            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
        )

    te = esc(_limpiar_vineta_literal(texto))

    return (
        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
        f"<w:r><w:t>{simbolo}</w:t></w:r>"
        f"<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>"
    )


def p_vineta_ejemplo(texto: str) -> str:
    if isinstance(texto, dict):
        clean = _limpiar_vineta_literal(rich_text(texto))
        runs = rich_runs(texto)
        if clean != rich_text(texto):
            runs = [{"text": clean}]
        return (
            f'    <w:p><w:pPr><w:pStyle w:val="Ejemplos-Vietanvl1"/></w:pPr>'
            f"<w:r><w:t>●</w:t></w:r>"
            f"<w:r><w:tab/></w:r>{_runs_to_xml(runs)}</w:p>"
        )

    te = esc(_limpiar_vineta_literal(texto))

    return (
        f'    <w:p><w:pPr><w:pStyle w:val="Ejemplos-Vietanvl1"/></w:pPr>'
        f"<w:r><w:t>●</w:t></w:r>"
        f"<w:r><w:tab/><w:t>{te}</w:t></w:r></w:p>"
    )


def p_vineta_bold(texto: str, nivel: int = 1) -> str:
    estilo = VINETA_EST.get(nivel, VINETA_EST[1])
    simbolo = esc(VINETA_SIM.get(nivel, VINETA_SIM[1]))
    te = esc(_limpiar_vineta_literal(texto))

    return (
        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
        f"<w:r><w:t>{simbolo}</w:t></w:r>"
        f"<w:r><w:tab/></w:r>"
        f"<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>{te}</w:t></w:r></w:p>"
    )


def p_desp(titulo: str, desc: str, nivel: int = 1) -> str:
    estilo = VINETA_EST.get(nivel, VINETA_EST[1])
    simbolo = esc(VINETA_SIM.get(nivel, VINETA_SIM[1]))
    titulo = titulo.rstrip(".: ")
    desc = desc.strip()
    desc = re.sub(r"::+", ":", desc)

    return (
        f'    <w:p><w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>'
        f"<w:r><w:t>{simbolo}</w:t></w:r>"
        f"<w:r><w:tab/></w:r>"
        f"<w:r><w:rPr><w:b/><w:bCs/></w:rPr>"
        f'<w:t xml:space="preserve">{esc(titulo)}. </w:t></w:r>'
        f"<w:r><w:t>{esc(desc)}</w:t></w:r></w:p>"
    )


def p_formula(texto: str) -> str:
    return p(texto, "Formula")


def p_opcion_test(letra: str, texto: str) -> str:
    texto = re.sub(r"^[a-h]\)\s*", "", texto.strip())
    return (
        '    <w:p><w:pPr><w:pStyle w:val="EjerciciosPregunta"/></w:pPr>'
        f"<w:r><w:t>{esc(letra)}) {esc(texto)}</w:t></w:r>"
        "</w:p>"
    )


def p_url_imagen(url: str) -> str:
    return p(rich_obj(url, [{"text": url, "link": url}]), "Cuerpoparrafo")


def p_url_recurso(url: str) -> str:
    return p(rich_obj(url, [{"text": url, "link": url}]), "Cuerpoparrafo")


def p_pie_imagen(texto: str) -> str:
    texto = texto.strip()
    texto = re.sub(r"^Pie de imagen:\s*", "", texto)

    return (
        '    <w:p><w:pPr><w:pStyle w:val="Cuerpoparrafo"/></w:pPr>'
        '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr>'
        '<w:t xml:space="preserve">Pie de imagen: </w:t></w:r>'
        f"<w:r><w:t>{esc(texto)}</w:t></w:r>"
        "</w:p>"
    )


def p_desc_imagen(texto: str) -> str:
    texto = texto.strip()
    texto = re.sub(r"^Descripci[oó]n de (la )?imagen:\s*", "", texto)

    return (
        '    <w:p><w:pPr><w:pStyle w:val="Cuerpoparrafo"/></w:pPr>'
        '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr>'
        '<w:t xml:space="preserve">Descripción de la imagen: </w:t></w:r>'
        f"<w:r><w:t>{esc(texto)}</w:t></w:r>"
        "</w:p>"
    )


# =============================================================================
# Interacciones
# =============================================================================

def parsear_interacciones(path: Path | None) -> dict[int, dict]:
    if not path or not path.exists():
        return {}

    if not _es_zip(path):
        return parsear_interacciones_texto(path)

    from docx import Document

    doc = Document(str(path))
    result: dict[int, dict] = {}

    for tbl in doc.tables:
        if not tbl.rows:
            continue

        header = tbl.rows[0].cells[0].text.strip() if tbl.rows[0].cells else ""
        m = RE_INTER.match(header)
        if not m:
            continue

        n = int(m.group(1))
        raw = tbl.rows[1].cells[0].text if len(tbl.rows) > 1 and tbl.rows[1].cells else ""
        raw = raw.replace("\r", "\n").strip()

        if raw.startswith("Opciones:"):
            result[n] = _parsear_interaccion_opciones(raw)
        elif "Desplegables:" in raw:
            result[n] = _parsear_interaccion_desplegables(raw)
        else:
            result[n] = {
                "tipo": "texto",
                "lineas": [_norm_line(x) for x in raw.splitlines() if _norm_line(x)]
            }

    return result


def parsear_interacciones_texto(path: Path) -> dict[int, dict]:
    text = path.read_text(encoding="utf-8", errors="replace")
    result: dict[int, dict] = {}

    parts = re.split(r"(?=Interacci[oó]n\s+\d+)", text)

    for part in parts:
        part = part.strip()
        if not part:
            continue

        lines = [re.sub(r"^\|?|\|?$", "", x).strip() for x in part.splitlines()]
        lines = [
            re.sub(r"\*+", "", x).strip()
            for x in lines
            if x.strip() and not re.match(r"^\|?\s*-+\s*\|?$", x)
        ]

        if not lines:
            continue

        m = RE_INTER.match(lines[0])
        if not m:
            continue

        n = int(m.group(1))
        raw = "\n".join(lines[1:]).strip()

        if raw.startswith("Opciones:"):
            result[n] = _parsear_interaccion_opciones(raw)
        elif "Desplegables:" in raw:
            result[n] = _parsear_interaccion_desplegables(raw)

    return result


def _parsear_interaccion_opciones(raw: str) -> dict:
    opciones: list[str] = []
    solucion: list[str] = []
    feedback: list[str] = []
    zona = "opciones"

    for line in raw.splitlines():
        line = line.strip()

        if not line:
            continue

        if line == "Opciones:":
            zona = "opciones"
            continue

        if re.match(r"^Soluci[oó]n:", line):
            zona = "solucion"
            rest = line.split(":", 1)[1].strip() if ":" in line else ""
            if rest:
                solucion.append(rest)
            continue

        if line.startswith("Feedback:"):
            zona = "feedback"
            rest = line.split(":", 1)[1].strip() if ":" in line else ""
            if rest:
                feedback.append(rest)
            continue

        if zona == "opciones":
            line = re.sub(r"^[a-h]\)\s*", "", line)
            opciones.append(line)
        elif zona == "solucion":
            solucion.append(line)
        elif zona == "feedback":
            feedback.append(line)

    return {
        "tipo": "opciones",
        "opciones": opciones,
        "solucion": " ".join(solucion).strip(),
        "feedback": " ".join(feedback).strip(),
    }


def _parsear_interaccion_desplegables(raw: str) -> dict:
    raw_items = raw.split("Desplegables:", 1)[1]
    lines = [x.rstrip() for x in raw_items.splitlines()]

    items = []
    i = 0

    while i < len(lines):
        while i < len(lines) and not lines[i].strip():
            i += 1

        if i >= len(lines):
            break

        titulo = lines[i].strip()
        i += 1
        contenido: list[str] = []

        while i < len(lines):
            current = lines[i].strip()

            if current and contenido and _parece_titulo_desplegable(current):
                nxt = _siguiente_no_vacia(lines, i + 1)
                if nxt is not None:
                    break

            contenido.append(lines[i])
            i += 1

        body, subitems = _separar_cuerpo_y_subitems(contenido)

        items.append({
            "titulo": titulo,
            "body": body,
            "subitems": subitems,
        })

    return {"tipo": "desplegables", "items": items}


def _parece_titulo_desplegable(line: str) -> bool:
    line = line.strip()

    if not line:
        return False
    if len(line) > 90:
        return False
    if line[-1] in ".;:":
        return False
    if line[0].islower():
        return False
    if re.match(
        r"^(Por ejemplo|Asimismo|Además|También|Puede|Se calcula|Número|Cuando|Una|Un consumidor|El cliente|La empresa)",
        line
    ):
        return False

    return True


def _siguiente_no_vacia(lines: list[str], start: int) -> str | None:
    for j in range(start, len(lines)):
        if lines[j].strip():
            return lines[j].strip()
    return None


def _separar_cuerpo_y_subitems(lines: list[str]) -> tuple[str, list[str]]:
    clean = [x.strip() for x in lines]

    while clean and not clean[0]:
        clean.pop(0)
    while clean and not clean[-1]:
        clean.pop()

    if not clean:
        return "", []

    # Si el contenido del desplegable empieza con una viñeta del online,
    # no lo convertimos en el cuerpo del título. Debe conservarse como
    # sublista; de lo contrario, el primer punto queda pegado al título
    # principal y los recursos de imagen pasan a viñetas.
    if re.match(r"^[·●○▪\-–]\s*", clean[0]):
        return "", [x for x in clean if x.strip()]

    # Si el desplegable contiene recursos de imagen, todo el contenido debe
    # conservarse como secuencia interna: subpuntos + URL + pie + descripción.
    # Así evitamos que el primer subpunto se fusione con el título principal
    # y que las URL/pies/descripciones se conviertan en viñetas.
    if any(RE_URL.match(x) or re.match(r"^Pie de imagen:", x, re.I) or re.match(r"^Descripci[oó]n de (la )?imagen:", x, re.I) for x in clean):
        return "", [x for x in clean if x.strip()]

    # Si hay varias líneas tipo "Etiqueta: explicación", son subpuntos del
    # desplegable, no un único párrafo unido.
    colon_lines = [x for x in clean if x.strip()]
    if len(colon_lines) > 1 and all(":" in x and len(x.split(":", 1)[0]) <= 70 for x in colon_lines):
        return "", colon_lines

    if "" not in clean:
        return _join_sin_doble_puntuacion(clean), []

    first_blank = clean.index("")
    body_lines = [x for x in clean[:first_blank] if x.strip()]
    rest = [x for x in clean[first_blank + 1:] if x.strip()]

    return _join_sin_doble_puntuacion(body_lines), rest


def _join_sin_doble_puntuacion(lines: list[str]) -> str:
    text = " ".join(x.strip() for x in lines if x.strip())
    text = re.sub(r"::+", ":", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    return text.strip()


def expandir_interaccion(n: int, interacciones: dict[int, dict]) -> list[dict]:
    inter = interacciones.get(n)

    if not inter:
        return []
    if inter.get("tipo") == "opciones":
        return []
    if inter.get("tipo") != "desplegables":
        return []

    bloques: list[dict] = []

    for item in inter.get("items", []):
        titulo = item.get("titulo", "").strip()
        body = item.get("body", "").strip()
        subitems = item.get("subitems", [])

        if not titulo or debe_elim(titulo):
            continue

        if body:
            bloques.append({
                "tipo": "desplegable_simple",
                "titulo": titulo,
                "contenido": body,
            })
        else:
            bloques.append({
                "tipo": "p_vineta_bold",
                "texto": titulo + ":",
                "nivel": 1,
            })

        for sub in subitems:
            sub = _limpiar_vineta_literal(sub)
            if not sub:
                continue

            # Los recursos gráficos y sus textos descriptivos no son elementos
            # de la enumeración: deben salir como URL/pie/descripción normales.
            if RE_URL.match(sub):
                bloques.append({
                    "tipo": "url_imagen" if _es_url_imagen(sub) else "url_recurso",
                    "texto": sub,
                })
            elif re.match(r"^Pie de imagen:", sub, re.I):
                bloques.append({"tipo": "pie_imagen", "texto": sub})
            elif re.match(r"^Descripci[oó]n de (la )?imagen:", sub, re.I):
                bloques.append({"tipo": "desc_imagen", "texto": sub})
            elif ":" in sub and len(sub.split(":", 1)[0]) <= 70:
                titulo_sub, desc_sub = sub.split(":", 1)
                bloques.append({
                    "tipo": "desplegable_simple_n2",
                    "titulo": titulo_sub.strip(),
                    "contenido": desc_sub.strip(),
                })
            else:
                bloques.append({
                    "tipo": "p_vineta",
                    "texto": sub,
                    "nivel": 2,
                })

    return bloques


# =============================================================================
# Extracción PDF / TXT
# =============================================================================

def extraer_texto_pdf(pdf: Path) -> list[str]:
    from pypdf import PdfReader
    reader = PdfReader(str(pdf))
    return [page.extract_text() or "" for page in reader.pages]


def parsear_pdf_o_texto(path: Path) -> dict:
    if path.suffix.lower() == ".pdf":
        paginas = extraer_texto_pdf(path)
        lines = []
        for pag in paginas:
            lines.extend(pag.splitlines())
    else:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()

    lines = [_norm_line(x) for x in lines]
    lines = [x for x in lines if x and not debe_elim(x) and not es_solo_pua(x)]

    est = {
        "titulo_unidad": "Unidad de aprendizaje 1",
        "titulo_modulo": "",
        "objetivos": [],
        "secciones": [],
    }

    current = None
    en_obj = False

    for line in lines:
        if line.startswith("Unidad de aprendizaje"):
            est["titulo_unidad"] = line
            continue

        if line.startswith("Los objetivos específicos"):
            en_obj = True
            continue

        if en_obj:
            if RE_SEC1.match(line) or line == "Introducción":
                en_obj = False
            else:
                est["objetivos"].append(_limpiar_vineta_literal(line))
                continue

        if not est["titulo_modulo"] and line and not RE_SEC1.match(line):
            if line != est["titulo_unidad"] and len(line) < 120:
                est["titulo_modulo"] = line
                continue

        m = RE_SEC1.match(line)
        if m:
            sec = {
                "num": m.group(1),
                "titulo": limpiar_titulo(m.group(2)),
                "bloques": [],
                "subsecciones": [],
            }
            est["secciones"].append(sec)
            current = sec
            continue

        if current:
            if RE_URL.match(line):
                current["bloques"].append({
                    "tipo": "url_imagen" if _es_url_imagen(line) else "url",
                    "url": line,
                })
            elif line.startswith("Pie de imagen:"):
                current["bloques"].append({
                    "tipo": "pie_imagen",
                    "texto": line[len("Pie de imagen:"):].strip(),
                })
            elif re.match(r"^Descripci[oó]n de (la )?imagen:", line):
                current["bloques"].append({
                    "tipo": "desc_imagen",
                    "texto": re.sub(r"^Descripci[oó]n de (la )?imagen:\s*", "", line),
                })
            else:
                current["bloques"].append({"tipo": "parrafo", "texto": line})

    return est


# =============================================================================
# Parser DOCX fuente
# =============================================================================

def parsear_docx_fuente(docx_path: Path, interacciones: dict[int, dict]) -> dict:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph as DParagraph
    from docx.table import Table as DTable

    doc = Document(str(docx_path))

    relmap = {}
    try:
        for rel in doc.part.rels.values():
            if str(rel.reltype).endswith('/hyperlink'):
                relmap[rel.rId] = rel.target_ref
    except Exception:
        relmap = {}

    def _runs_desde_para(para) -> list[dict]:
        runs: list[dict] = []

        def add_run_xml(r_el, link_target: str | None = None):
            texts = []
            for t_el in r_el.findall('.//' + qn('w:t')):
                texts.append(t_el.text or '')
            if not texts:
                return
            text = ''.join(texts)
            rpr = r_el.find(qn('w:rPr'))
            bold_r = bool(rpr is not None and rpr.find(qn('w:b')) is not None)
            italic_r = bool(rpr is not None and rpr.find(qn('w:i')) is not None)
            item = {'text': text}
            if bold_r:
                item['bold'] = True
            if italic_r:
                item['italic'] = True
            if link_target:
                item['link'] = link_target
            runs.append(item)

        for child_el in para._p.iterchildren():
            if child_el.tag == qn('w:r'):
                add_run_xml(child_el)
            elif child_el.tag == qn('w:hyperlink'):
                rid = child_el.get(qn('r:id'))
                target = relmap.get(rid, '')
                for r_el in child_el.findall(qn('w:r')):
                    add_run_xml(r_el, target)

        if not runs and para.text:
            runs = [{'text': para.text}]
        return _trim_runs(runs)

    def _rich_para(para) -> dict:
        runs = _materializar_hipervinculos_para_papel(_runs_desde_para(para))
        texto = _norm_line(''.join(r.get('text', '') for r in runs) if runs else para.text)
        return rich_obj(texto, runs or [{'text': texto}])

    est = {
        "titulo_unidad": "",
        "titulo_modulo": "",
        "objetivos": [],
        "secciones": [],
    }

    current_sec = None
    current_sub = None
    current_sub2 = None
    sec_count = 0
    en_objetivos = False
    blk = None

    def activos() -> list:
        if current_sub2:
            return current_sub2["bloques"]
        if current_sub:
            return current_sub["bloques"]
        if current_sec:
            return current_sec["bloques"]
        return []

    def cerrar_ejercicio_pendiente():
        nonlocal blk

        if not blk or blk.get("tipo") != "_ejercicio_pendiente":
            return False

        inter = blk.get("inter", {})
        letras = "abcdefgh"
        opciones = []

        for j, opt in enumerate(inter.get("opciones", [])):
            opciones.append({
                "letra": letras[j],
                "texto": re.sub(r"^[a-h]\)\s*", "", opt.strip()),
            })

        activos().append({
            "tipo": "tarea",
            "etiqueta": blk.get("etiqueta", "Actividad 1"),
            "lineas": [x for x in blk.get("lineas", []) if rich_text(x).strip()],
            "opciones": opciones,
            "_solucion": inter.get("solucion", ""),
            "_feedback": inter.get("feedback", ""),
        })

        blk = None
        return True

    def flush():
        nonlocal blk

        if not blk:
            return

        if blk.get("tipo") == "_ejercicio_pendiente":
            cerrar_ejercicio_pendiente()
            return

        if blk.get("lineas") or blk.get("tipo") in {
            "hilo_conductor", "video", "ejemplo", "importante",
            "definicion", "sabias_que",
        }:
            b = {k: v for k, v in blk.items() if not k.startswith("_")}
            activos().append(b)

        blk = None

    def add(b: dict):
        flush()
        activos().append(b)

    def nueva_sec(titulo: str):
        nonlocal current_sec, current_sub, current_sub2, sec_count

        flush()
        sec_count += 1
        current_sec = {
            "num": str(sec_count),
            "titulo": limpiar_titulo(titulo),
            "bloques": [],
            "subsecciones": [],
        }
        est["secciones"].append(current_sec)
        current_sub = None
        current_sub2 = None

    def nueva_sub(num: str, titulo: str):
        nonlocal current_sub, current_sub2

        flush()

        if not current_sec:
            nueva_sec("Introducción")

        current_sub = {
            "num": num,
            "titulo": limpiar_titulo(titulo),
            "bloques": [],
            "subsecciones": [],
        }
        current_sec["subsecciones"].append(current_sub)
        current_sub2 = None

    def nueva_sub2(num: str, titulo: str):
        nonlocal current_sub2

        flush()

        if not current_sub:
            return

        current_sub2 = {
            "num": num,
            "titulo": limpiar_titulo(titulo),
            "bloques": [],
            "subsecciones": [],
        }
        current_sub["subsecciones"].append(current_sub2)

    special_styles = {
        "Hilo conductor": "hilo_conductor",
        "Ejemplo": "ejemplo",
        "Vídeo": "video",
        "Actividad colaborativa": "actividad_complementaria",
        "Importante": "importante",
        "Aplicación práctica": "tarea",
        "Sabiasque": "sabias_que",
        "Definición": "definicion",
    }

    list_styles = {
        "List Paragraph",
        "Viñeta nvl1 1d",
        "Viñeta nvl2 1d",
        "Vietanvl11d",
        "Vietanvl21d",
        "Ejemplos - Viñeta nvl1",
        "Ejemplos-Vietanvl1",
        "Ejemplos Vieta nvl1",
    }

    for child in doc.element.body:
        is_para = child.tag == qn("w:p")
        is_tbl = child.tag == qn("w:tbl")

        if is_tbl:
            tbl = DTable(child, doc)
            flush()

            for row in tbl.rows:
                txt = row.cells[0].text.strip() if row.cells else ""
                if txt and current_sec:
                    activos().append({
                        "tipo": "p_vineta",
                        "texto": txt,
                        "nivel": 1,
                    })
            continue

        if not is_para:
            continue

        para = DParagraph(child, doc)
        style = _style_name(para)
        txt = _norm_line(para.text)
        rich = _rich_para(para)
        bold = _is_bold(para)

        if es_solo_pua(txt):
            continue

        if txt in {"", "Cambio de pantalla", "Específicos", "Objetivos"} and style not in special_styles:
            if txt == "Objetivos":
                en_objetivos = True
            continue

        if debe_elim(txt):
            continue

        if blk and blk.get("tipo") == "_ejercicio_pendiente":
            es_fin = False

            if style.startswith("Heading"):
                es_fin = True
            elif RE_INTER.match(txt) and bold:
                es_fin = True
            elif txt in BLOQUES_ESP:
                es_fin = True
            elif txt.startswith(("Duración:", "Objetivo:")):
                continue

            if es_fin:
                cerrar_ejercicio_pendiente()
            else:
                item = rich
                if txt.startswith("Enunciado:"):
                    stripped = txt[len("Enunciado:"):].strip()
                    item = rich_obj(stripped, [{"text": stripped}])
                if txt and not txt.startswith(("Solución:", "Feedback:")):
                    blk["lineas"].append(item)
                continue

        if style == "Title" or style.startswith("_TITULO UNIDAD"):
            m = re.match(r"Unidad de aprendizaje\s+(\d+)(?:[.\-–\s]+(.+))?", txt, re.I)

            if m:
                est["titulo_unidad"] = f"Unidad de aprendizaje {m.group(1)}"
                if m.group(2):
                    est["titulo_modulo"] = m.group(2).strip()
            elif est["titulo_unidad"] and not est["titulo_modulo"]:
                est["titulo_modulo"] = txt
            continue

        if not est["titulo_unidad"]:
            m = re.match(r"Unidad de aprendizaje\s+(\d+)", txt, re.I)
            if m:
                est["titulo_unidad"] = f"Unidad de aprendizaje {m.group(1)}"
                continue

        if est["titulo_unidad"] and not est["titulo_modulo"] and txt and not txt.startswith("Los objetivos"):
            if style in {"Title", "_TITULO UNIDAD 2"} or (not RE_SEC1.match(txt) and len(txt) < 100):
                est["titulo_modulo"] = txt
                continue

        if txt.startswith("Los objetivos específicos"):
            en_objetivos = True
            continue

        if en_objetivos:
            if style.startswith("Heading") or RE_SEC1.match(txt) or txt == "Introducción":
                en_objetivos = False
            elif txt and not txt.startswith("CE ") and not re.match(r"^[a-e]\) Se han", txt):
                est["objetivos"].append(_limpiar_vineta_rich(rich))
                continue

        if style == "Heading 1" or style == "1 Título nvl1":
            if txt == "Objetivos":
                en_objetivos = True
                continue

            if txt == "Introducción":
                nueva_sec("Introducción")
                continue

            m = RE_SEC1.match(txt)
            if m:
                nueva_sec(m.group(2))
            else:
                nueva_sec(txt)
            continue

        if style == "Heading 2" or style == "2 Título nvl2":
            m = RE_SEC2.match(txt)
            if m:
                nueva_sub(f"{m.group(1)}.{m.group(2)}", m.group(3))
            continue

        if style == "Heading 3" or style == "3 Título nvl3":
            m = RE_SEC3.match(txt)
            if m:
                nueva_sub2(f"{m.group(1)}.{m.group(2)}.{m.group(3)}", m.group(4))
            continue

        if current_sec is None and txt == "Introducción":
            nueva_sec("Introducción")
            continue

        # Dentro de una tarea de evaluación, las líneas numeradas del enunciado
        # (1., 2., 3...) son instrucciones, no títulos de nuevas secciones.
        if blk and blk.get("_estilo") == "Aplicación práctica" and txt:
            if not txt.startswith(("Duración:", "Objetivo:", "Enunciado:")) and not debe_elim(txt):
                blk.setdefault("lineas", []).append(rich)
            continue

        m3 = RE_SEC3.match(txt)
        if m3 and current_sec:
            nueva_sub2(f"{m3.group(1)}.{m3.group(2)}.{m3.group(3)}", m3.group(4))
            continue

        m2 = RE_SEC2.match(txt)
        if m2 and current_sec:
            nueva_sub(f"{m2.group(1)}.{m2.group(2)}", m2.group(3))
            continue

        m1 = RE_SEC1.match(txt)
        if m1 and len(m1.group(2)) < 120:
            nueva_sec(m1.group(2))
            continue

        if not current_sec:
            continue

        mi = RE_INTER.match(txt)
        if mi and bold:
            flush()

            n = int(mi.group(1))
            label = (mi.group(2) or "").strip()
            inter = interacciones.get(n, {})

            if inter.get("tipo") == "opciones":
                mt = re.search(r"Actividad de evaluaci[oó]n\s+(\d+)", label, re.I)
                num = mt.group(1) if mt else "1"

                blk = {
                    "tipo": "_ejercicio_pendiente",
                    "_estilo": "_ejercicio",
                    "etiqueta": f"Actividad {num}",
                    "lineas": [],
                    "inter": inter,
                }
            else:
                for b in expandir_interaccion(n, interacciones):
                    activos().append(b)
            continue

        if style in special_styles:
            if style == "Sabiasque":
                if txt in {"Sabías que…", "Sabías que...", "Definición", "Importante"}:
                    flush()
                    tipo_real = {
                        "Sabías que…": "sabias_que",
                        "Sabías que...": "sabias_que",
                        "Definición": "definicion",
                        "Importante": "importante",
                    }.get(txt, "sabias_que")

                    blk = {
                        "tipo": tipo_real,
                        "etiqueta": txt,
                        "lineas": [],
                        "_estilo": style,
                    }
                else:
                    if not blk or blk.get("_estilo") != style:
                        flush()
                        blk = {
                            "tipo": "sabias_que",
                            "etiqueta": "Sabías que…",
                            "lineas": [],
                            "_estilo": style,
                        }
                    blk["lineas"].append(_limpiar_vineta_rich(rich))
                continue

            if style == "Definición":
                if txt == "Definición":
                    flush()
                    blk = {
                        "tipo": "definicion",
                        "etiqueta": "Definición",
                        "lineas": [],
                        "_estilo": style,
                    }
                elif blk and blk.get("_estilo") == style:
                    blk["lineas"].append(rich)
                else:
                    flush()
                    blk = {
                        "tipo": "definicion",
                        "etiqueta": "Definición",
                        "lineas": [rich],
                        "_estilo": style,
                    }
                continue

            if style == "Hilo conductor":
                if txt == "Hilo conductor":
                    flush()
                    blk = {
                        "tipo": "hilo_conductor",
                        "etiqueta": "Hilo conductor",
                        "lineas": [],
                        "_estilo": style,
                    }
                elif blk and blk.get("_estilo") == style:
                    blk["lineas"].append(rich)
                continue

            if style == "Ejemplo":
                if txt == "Ejemplo":
                    flush()
                    blk = {
                        "tipo": "ejemplo",
                        "etiqueta": "Ejemplo",
                        "lineas": [],
                        "_estilo": style,
                    }
                elif blk and blk.get("_estilo") == style:
                    blk["lineas"].append(rich)
                continue

            if style == "Vídeo":
                if txt == "Vídeo":
                    flush()
                    blk = {
                        "tipo": "video",
                        "etiqueta": "Vídeo",
                        "lineas": [],
                        "_estilo": style,
                    }
                elif blk and blk.get("_estilo") == style:
                    blk["lineas"].append(rich)
                continue

            if style == "Importante":
                if txt == "Importante" or not blk or blk.get("_estilo") != style:
                    flush()
                    blk = {
                        "tipo": "importante",
                        "etiqueta": "Importante",
                        "lineas": [],
                        "_estilo": style,
                    }
                    if txt and txt != "Importante":
                        blk["lineas"].append(rich)
                else:
                    blk["lineas"].append(rich)
                continue

            if style == "Actividad colaborativa":
                if not blk or blk.get("_estilo") != style:
                    flush()
                    blk = {
                        "tipo": "actividad_complementaria",
                        "etiqueta": "Actividad complementaria",
                        "lineas": [],
                        "_estilo": style,
                    }

                if txt and not txt.lower().startswith("actividad"):
                    # v10: se respeta la numeración visible del online (por ejemplo, "1.").
                    blk["lineas"].append(rich)
                continue

            if style == "Aplicación práctica":
                if not blk or blk.get("_estilo") != style:
                    flush()
                    mt = re.search(r"Tarea de evaluaci[oó]n\s+(\d+)", txt, re.I)
                    num = mt.group(1) if mt else "1"
                    blk = {
                        "tipo": "tarea",
                        "etiqueta": f"Tarea {num}",
                        "lineas": [],
                        "_estilo": style,
                    }
                else:
                    if txt and not txt.startswith(("Duración:", "Objetivo:", "Enunciado:")):
                        blk["lineas"].append(rich)
                continue

        if blk and blk.get("_estilo") and style in list_styles:
            if txt:
                blk.setdefault("lineas", []).append(_prefix_rich("● ", _limpiar_vineta_rich(rich)))
            continue

        if blk and blk.get("_estilo") and style not in special_styles:
            flush()

        if txt in BLOQUES_ESP:
            flush()

            tipo_map = {
                "Nota": "nota",
                "Ejemplo": "ejemplo",
                "Sabías que...": "sabias_que",
                "Sabías que…": "sabias_que",
                "Consejo": "consejo",
                "Definición": "definicion",
                "Hilo conductor": "hilo_conductor",
                "Para saber más": "para_saber_mas",
                "Vídeo": "video",
                "Importante": "importante",
                "Recuerda": "recuerda",
            }

            blk = {
                "tipo": tipo_map.get(txt, "ejemplo"),
                "etiqueta": txt,
                "lineas": [],
                "_estilo": "_texto",
            }
            continue

        if RE_URL.match(txt):
            add({
                "tipo": "url_imagen" if _es_url_imagen(txt) else "url",
                "url": txt,
            })
            continue

        if txt.startswith("Pie de imagen:"):
            add({
                "tipo": "pie_imagen",
                "texto": txt[len("Pie de imagen:"):].strip(),
            })
            continue

        if re.match(r"^Descripci[oó]n de (la )?imagen:", txt):
            add({
                "tipo": "desc_imagen",
                "texto": re.sub(r"^Descripci[oó]n de (la )?imagen:\s*", "", txt),
            })
            continue

        if style in list_styles:
            nivel = 2 if "nvl2" in style.lower() or "21" in style else 1
            add({
                "tipo": "p_vineta",
                "texto": rich,
                "nivel": nivel,
            })
            continue

        mo = RE_OPCION.match(txt)
        if mo:
            add({
                "tipo": "opcion_test_suelta",
                "letra": mo.group(1),
                "texto": mo.group(2),
            })
            continue

        add({
            "tipo": "parrafo",
            "texto": rich,
        })

    flush()

    if not est["titulo_unidad"]:
        est["titulo_unidad"] = "Unidad de aprendizaje 1"

    if not est["titulo_modulo"]:
        est["titulo_modulo"] = ""

    return est


# =============================================================================
# XML de bloques
# =============================================================================

def bloques_xml(bloques: list[dict]) -> list[str]:
    out: list[str] = []

    for b in bloques:
        t = b.get("tipo", "")

        if t == "parrafo":
            texto = b.get("texto", "")
            texto_plano = rich_text(texto)
            if RE_FORMULA.search(texto_plano) and len(texto_plano) < 200:
                out.append(p_formula(texto))
            else:
                out.append(p(texto, "Cuerpoparrafo"))

        elif t == "p_vineta":
            out.append(p_vineta(b.get("texto", ""), b.get("nivel", 1)))

        elif t == "p_vineta_bold":
            out.append(p_vineta_bold(b.get("texto", ""), b.get("nivel", 1)))

        elif t == "p_vineta_ejemplo":
            out.append(p_vineta_ejemplo(b.get("texto", "")))

        elif t == "desplegable":
            titulo = b.get("titulo", "")
            desc = b.get("descripcion", "")
            if desc:
                out.append(p_desp(titulo, desc))
            else:
                out.append(p_vineta_bold(titulo))

        elif t == "desplegable_simple":
            out.append(p_desp(b.get("titulo", ""), b.get("contenido", "")))

        elif t == "desplegable_simple_n2":
            out.append(p_desp(b.get("titulo", ""), b.get("contenido", ""), 2))

        elif t == "url_imagen":
            out.append(p_url_imagen(b.get("texto", "")))

        elif t == "url_recurso":
            out.append(p_url_recurso(b.get("texto", "")))

        elif t == "pie_imagen":
            out.append(p_pie_imagen(b.get("texto", "")))

        elif t == "desc_imagen":
            out.append(p_desc_imagen(b.get("texto", "")))

        elif t == "desplegable_multi":
            out.append(p_vineta_bold(b.get("titulo", "").rstrip(":") + ":"))
            for item in b.get("items", []):
                out.append(p_vineta(item, 2))

        elif t in {
            "nota", "ejemplo", "sabias_que", "consejo", "definicion",
            "hilo_conductor", "para_saber_mas", "video", "importante",
        }:
            out.append(p(b.get("etiqueta", ""), "Ejemplos-01lneainicio"))

            modo_lista = False

            for line in b.get("lineas", []):
                line_txt = rich_text(line).strip()

                if not line_txt:
                    continue

                if RE_URL.match(line_txt):
                    out.append(p_url_recurso(line_txt))
                    modo_lista = False
                    continue

                if _parece_item_lista_en_bloque(line_txt, modo_lista):
                    out.append(p_vineta_ejemplo(line if isinstance(line, dict) else _limpiar_vineta_literal(line_txt)))
                    continue

                out.append(p(line, "Ejemplos-Cuerpoparrafo"))
                modo_lista = _abre_modo_lista(line_txt)

            out.append(p("", "Ejemplos-02lneafin"))

        elif t == "recuerda":
            out.append(p(b.get("etiqueta", "Recuerda"), "Recuerda-00lneainicio"))

            for line in b.get("lineas", []):
                if rich_text(line).strip():
                    out.append(p(line, "Recuerda-cuerpoparrafo"))

            out.append(p("", "Recuerda-01lneafin"))

        elif t == "tarea":
            out.append(p(b.get("etiqueta", "Tarea"), "Ejemplos-01lneainicio"))

            for line in b.get("lineas", []):
                if rich_text(line).strip():
                    out.append(p(line, "EjerciciosPregunta"))

            for opt in b.get("opciones", []):
                out.append(p_opcion_test(opt.get("letra", ""), opt.get("texto", "")))

            out.append(p("", "Ejemplos-02lneafin"))

        elif t == "actividad_complementaria":
            out.append(p(b.get("etiqueta", "Actividad complementaria"), "Ejemplos-01lneainicio"))

            for line in b.get("lineas", []):
                if rich_text(line).strip():
                    out.append(p(line, "EjerciciosPregunta"))

            out.append(p("", "Ejemplos-02lneafin"))

        elif t == "url_imagen":
            out.append(p_url_imagen(b.get("url", "")))

        elif t == "url":
            out.append(p_url_recurso(b.get("url", "")))

        elif t == "pie_imagen":
            out.append(p_pie_imagen(b.get("texto", "")))

        elif t == "desc_imagen":
            out.append(p_desc_imagen(b.get("texto", "")))

        elif t == "imagen":
            if b.get("url"):
                out.append(p_url_imagen(b["url"]))
            if b.get("pie"):
                out.append(p_pie_imagen(b["pie"]))
            if b.get("descripcion"):
                out.append(p_desc_imagen(b["descripcion"]))

        elif t == "opcion_test_suelta":
            out.append(p_opcion_test(b.get("letra", ""), b.get("texto", "")))

        elif t == "parrafo_formula":
            out.append(p_formula(b.get("texto", "")))

    return out


# =============================================================================
# Gráficos / imágenes / SmartArt desde el DOCX de referencia
# =============================================================================

def _leer_document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(str(docx_path), "r") as z:
        return z.read("word/document.xml").decode("utf-8")


def _extraer_body_xml(document_xml: str) -> str:
    m = re.search(r"<w:body>([\s\S]*?)</w:body>", document_xml)
    return m.group(1) if m else ""


def _extraer_parrafos_body(body_xml: str) -> list[str]:
    """
    IMPORTANTE:
    No usar <w:p[\\s\\S]*?</w:p> porque también casa <w:pPr>.
    """
    return re.findall(r"<w:p(?:\s[^>]*)?>[\s\S]*?</w:p>", body_xml)


def _extraer_document_attrs(docx_path: Path) -> str:
    """
    Usa los namespace reales del documento de referencia.
    Esto evita perder gráficos por faltar xmlns:wp14, wpc, wpg, wps, etc.
    """
    fallback = (
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'mc:Ignorable="w14 w15 wp14"'
    )

    if not docx_path.exists() or not _es_zip(docx_path):
        return fallback

    try:
        xml = _leer_document_xml(docx_path)
    except Exception:
        return fallback

    m = re.search(r"<w:document\s+([^>]*)>", xml)
    if not m:
        return fallback

    attrs = m.group(1)

    needed = {
        "xmlns:wp14": 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
        "xmlns:a": 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        "xmlns:pic": 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
        "xmlns:dgm": 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"',
    }

    for key, val in needed.items():
        if key not in attrs:
            attrs += " " + val

    if "mc:Ignorable" not in attrs:
        attrs += ' mc:Ignorable="w14 w15 wp14"'

    return attrs


def _texto_plano_parrafo_xml(par_xml: str) -> str:
    textos = re.findall(r"<w:t[^>]*>(.*?)</w:t>", par_xml)
    txt = " ".join(textos)
    txt = re.sub(r"<[^>]+>", "", txt)
    txt = (
        txt.replace("&amp;", "&")
           .replace("&lt;", "<")
           .replace("&gt;", ">")
           .replace("&quot;", '"')
    )
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def _normalizar_contexto(txt: str) -> str:
    txt = txt or ""
    txt = txt.lower()
    txt = (
        txt.replace("á", "a")
           .replace("é", "e")
           .replace("í", "i")
           .replace("ó", "o")
           .replace("ú", "u")
           .replace("ü", "u")
           .replace("ñ", "n")
    )
    txt = re.sub(r"\s+", " ", txt)
    txt = re.sub(r"[^\w ]+", "", txt)
    return txt.strip()


def _parrafo_tiene_grafico(par_xml: str) -> bool:
    """
    Detecta cualquier gráfico insertado en Word:
      - imagen normal: <a:blip r:embed="...">
      - SmartArt / diagrama: <dgm:relIds ...>
      - formas / dibujos: <w:drawing>
      - VML antiguo: <w:pict>
    """
    return (
        "<w:drawing" in par_xml
        or "<w:pict" in par_xml
        or "<a:graphicData" in par_xml
        or "<dgm:relIds" in par_xml
        or "<a:blip" in par_xml
        or "r:embed=" in par_xml
        or "r:link=" in par_xml
    )


def extraer_graficos_con_contexto(docx_path: Path) -> list[dict]:
    """
    Extrae todos los gráficos del DOCX de referencia:
      - imágenes normales
      - SmartArt
      - diagramas
      - formas
      - dibujos

    Copia el párrafo <w:drawing> completo.
    """
    if not docx_path.exists() or not _es_zip(docx_path):
        return []

    try:
        document_xml = _leer_document_xml(docx_path)
    except Exception:
        return []

    body = _extraer_body_xml(document_xml)
    if not body:
        return []

    paras = _extraer_parrafos_body(body)

    graficos = []

    for i, par in enumerate(paras):
        if not _parrafo_tiene_grafico(par):
            continue

        prev_txt = ""
        next_txt = ""

        for j in range(i - 1, -1, -1):
            t = _texto_plano_parrafo_xml(paras[j])
            if t:
                prev_txt = t
                break

        for j in range(i + 1, len(paras)):
            t = _texto_plano_parrafo_xml(paras[j])
            if t:
                next_txt = t
                break

        graficos.append({
            "xml": "    " + par,
            "prev": prev_txt,
            "next": next_txt,
            "prev_norm": _normalizar_contexto(prev_txt),
            "next_norm": _normalizar_contexto(next_txt),
            "insertado": False,
        })

    return graficos


def insertar_graficos_por_contexto(pars: list[str], graficos: list[dict]) -> list[str]:
    """
    Recoloca gráficos del ejemplo dentro del documento generado.

    Prioridad:
      1. Insertar antes del texto posterior.
      2. Insertar después del texto anterior.
      3. Insertar después del título Resumen.
      4. Insertar al final.
    """
    if not graficos:
        return pars

    resultado = []

    for par in pars:
        par_txt = _texto_plano_parrafo_xml(par)
        par_norm = _normalizar_contexto(par_txt)

        for g in graficos:
            if g["insertado"]:
                continue

            next_norm = g.get("next_norm", "")
            if not next_norm:
                continue

            claves = [
                next_norm[:220],
                next_norm[:180],
                next_norm[:140],
                next_norm[:100],
                next_norm[:70],
                next_norm[:45],
            ]

            if any(k and (par_norm.startswith(k) or k in par_norm) for k in claves):
                resultado.append(g["xml"])
                g["insertado"] = True

        resultado.append(par)

        for g in graficos:
            if g["insertado"]:
                continue

            prev_norm = g.get("prev_norm", "")
            if not prev_norm:
                continue

            claves = [
                prev_norm[:220],
                prev_norm[:180],
                prev_norm[:140],
                prev_norm[:100],
                prev_norm[:70],
                prev_norm[:45],
            ]

            if any(k and (par_norm.startswith(k) or k in par_norm) for k in claves):
                resultado.append(g["xml"])
                g["insertado"] = True

    if any(not g["insertado"] for g in graficos):
        nuevo = []
        metidos_en_resumen = False

        for par in resultado:
            nuevo.append(par)

            txt = _normalizar_contexto(_texto_plano_parrafo_xml(par))
            if not metidos_en_resumen and re.search(r"\b\d+\s+resumen\b|\bresumen\b", txt):
                for g in graficos:
                    if not g["insertado"]:
                        nuevo.append(g["xml"])
                        g["insertado"] = True
                metidos_en_resumen = True

        resultado = nuevo

    for g in graficos:
        if not g["insertado"]:
            resultado.append(g["xml"])
            g["insertado"] = True

    return resultado


# =============================================================================
# Generación DOCX
# =============================================================================

def extraer_sectpr(docx_path: Path) -> str:
    if not docx_path.exists() or not _es_zip(docx_path):
        return "<w:sectPr/>"

    try:
        xml = _leer_document_xml(docx_path)
    except Exception:
        return "<w:sectPr/>"

    matches = re.findall(r"<w:sectPr[\s\S]*?</w:sectPr>", xml)
    if matches:
        return matches[-1]

    m = re.search(r"<w:sectPr[^>]*/>", xml)
    if m:
        return m.group(0)

    return "<w:sectPr/>"


def crear_docx_minimo() -> dict[str, bytes]:
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        "</Types>"
    ).encode("utf-8")

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    word_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        '<w:rPr><w:sz w:val="24"/></w:rPr>'
        "</w:style>"
        "</w:styles>"
    ).encode("utf-8")

    return {
        "[Content_Types].xml": content_types,
        "_rels/.rels": rels,
        "word/_rels/document.xml.rels": word_rels,
        "word/styles.xml": styles,
    }


def _cargar_paquete_base(base: Path | None, ejemplo: Path) -> dict[str, bytes]:
    """
    Carga el paquete base.

    Clave:
      - Después copia TODO el paquete del ejemplo excepto word/document.xml.
      - Así los rId de SmartArt, diagramas y dibujos siguen existiendo.
    """
    archivos: dict[str, bytes] = {}

    if base and base.exists() and _es_zip(base):
        with zipfile.ZipFile(str(base), "r") as zin:
            archivos = {name: zin.read(name) for name in zin.namelist()}
    else:
        archivos = crear_docx_minimo()

    if ejemplo and ejemplo.exists() and _es_zip(ejemplo):
        with zipfile.ZipFile(str(ejemplo), "r") as zej:
            for name in zej.namelist():
                if name == "word/document.xml":
                    continue
                archivos[name] = zej.read(name)

    return archivos


def generar_docx(est: dict, ejemplo: Path, plantilla: Path, salida: Path):
    _reset_hyperlinks_out()
    ns = _extraer_document_attrs(ejemplo)

    pars: list[str] = []

    pars.append(p(est.get("titulo_unidad", "Unidad de aprendizaje 1"), "TITULOUNIDAD1"))
    pars.append(p(est.get("titulo_modulo", ""), "TITULOUNIDAD2", negrita=True))

    objetivos = est.get("objetivos", [])

    if objetivos:
        pars.append(p("Los objetivos específicos de esta Unidad de Aprendizaje son:", "Cuerpoparrafo"))
        for obj in objetivos:
            pars.append(p_vineta(obj, 1))

    for sec in est.get("secciones", []):
        pars.append(p(f'{sec.get("num", "")}. {sec.get("titulo", "")}', "1Ttulonvl1"))
        pars.extend(bloques_xml(sec.get("bloques", [])))

        for sub in sec.get("subsecciones", []):
            pars.append(p(f'{sub.get("num", "")} {sub.get("titulo", "")}', "2Ttulonvl2"))
            pars.extend(bloques_xml(sub.get("bloques", [])))

            for sub2 in sub.get("subsecciones", []):
                pars.append(p(f'{sub2.get("num", "")} {sub2.get("titulo", "")}', "3Ttulonvl3"))
                pars.extend(bloques_xml(sub2.get("bloques", [])))

    graficos = extraer_graficos_con_contexto(ejemplo)
    print(f"  Gráficos encontrados en referencia: {len(graficos)}")
    pars = insertar_graficos_por_contexto(pars, graficos)
    print(f"  Gráficos insertados: {sum(1 for g in graficos if g.get('insertado'))}")

    sectpr = extraer_sectpr(ejemplo)

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<w:document {ns}><w:body>\n"
        + "\n".join(pars)
        + "\n"
        + sectpr
        + "\n</w:body></w:document>"
    ).encode("utf-8")

    base = plantilla if plantilla and plantilla.exists() and _es_zip(plantilla) else ejemplo
    archivos = _cargar_paquete_base(base, ejemplo)
    archivos["word/document.xml"] = document_xml
    if "word/_rels/document.xml.rels" in archivos:
        archivos["word/_rels/document.xml.rels"] = _patch_document_rels_hyperlinks(archivos["word/_rels/document.xml.rels"])

    salida.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(salida), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)


# =============================================================================
# Main
# =============================================================================

def buscar_interacciones_auto(unidad: Path) -> Path | None:
    candidatos = sorted(unidad.parent.glob("interacciones_*.docx"))
    return candidatos[0] if candidatos else None


def main():
    args = sys.argv[1:]
    inter_path = None

    if len(args) == 2:
        unidad, ejemplo = (Path(x) for x in args)
        plantilla = ejemplo
        salida = unidad.with_name(unidad.stem + "_resultado.docx")

    elif len(args) == 3:
        unidad, ejemplo, salida = (Path(x) for x in args)
        plantilla = ejemplo

    elif len(args) == 4:
        unidad = Path(args[0])
        ejemplo = Path(args[1])
        inter_path = Path(args[2])
        salida = Path(args[3])
        plantilla = ejemplo

    elif len(args) == 5:
        unidad = Path(args[0])
        ejemplo = Path(args[1])
        plantilla = Path(args[2])
        inter_path = Path(args[3])
        salida = Path(args[4])

    else:
        print("Uso:")
        print("  python conversor_papel.py UNIDAD.docx EJEMPLO.docx SALIDA.docx")
        print("  python conversor_papel.py UNIDAD.docx EJEMPLO.docx INTERACCIONES.docx SALIDA.docx")
        print("  python conversor_papel.py UNIDAD.docx EJEMPLO.docx PLANTILLA.docx INTERACCIONES.docx SALIDA.docx")
        sys.exit(1)

    for path, label in [(unidad, "Unidad"), (ejemplo, "Ejemplo maquetado")]:
        if not path.exists():
            print(f"ERROR: no existe {label}: {path}")
            sys.exit(1)

    if inter_path is None and unidad.suffix.lower() in {".docx", ".doc"}:
        inter_path = buscar_interacciones_auto(unidad)
        if inter_path:
            print(f"→ Interacciones detectadas: {inter_path.name}")

    interacciones = {}

    if inter_path and inter_path.exists():
        print(f"→ Parseando interacciones: {inter_path.name}")
        interacciones = parsear_interacciones(inter_path)
        print(f"  {len(interacciones)} interacciones cargadas")

    print(f"→ Parseando unidad: {unidad.name}")

    if unidad.suffix.lower() in {".docx", ".doc"} and _es_zip(unidad):
        est = parsear_docx_fuente(unidad, interacciones)
    else:
        est = parsear_pdf_o_texto(unidad)

    print(f'  {est.get("titulo_unidad", "")} — {est.get("titulo_modulo", "")}')
    print(f'  Objetivos: {len(est.get("objetivos", []))}')
    print(f'  Secciones: {len(est.get("secciones", []))}')

    total_bloques = 0

    for sec in est.get("secciones", []):
        nb = len(sec.get("bloques", []))
        ns = len(sec.get("subsecciones", []))
        nsb = sum(len(s.get("bloques", [])) for s in sec.get("subsecciones", []))
        total_bloques += nb + nsb
        print(f'    {sec.get("num")}. {sec.get("titulo")} [{nb} bloques, {ns} subsecs, {nsb} bloques subsec]')

    print(f"  Total bloques: {total_bloques}")

    print(f"→ Generando salida: {salida.name}")
    generar_docx(est, ejemplo, plantilla, salida)
    print(f"✓ Listo: {salida}")


if __name__ == "__main__":
    main()
