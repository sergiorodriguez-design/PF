import base64
import io
import os
import re
import subprocess
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

from flask import Flask, request, jsonify

app = Flask(__name__)

CONVERSOR = Path(__file__).parent / "conversor_papel.py"


# ── Health ────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
def home():
    return jsonify({"ok": True, "service": "conversor-papel", "status": "online"})


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True})


# ── /convertir  (una sola unidad) ─────────────────────────────────────────────

@app.route("/convertir", methods=["POST"])
def convertir():
    try:
        data = request.get_json(force=True)

        nombre_salida = data.get("nombreSalida") or "unidad_papel.docx"
        archivos = data.get("archivos", {})

        unidad     = archivos.get("unidad")
        plantilla  = archivos.get("plantilla")
        interacc   = archivos.get("interacciones")

        if not unidad:
            return jsonify({"ok": False, "error": "Falta archivo de unidad."}), 400
        if not plantilla:
            return jsonify({"ok": False, "error": "Falta archivo de plantilla."}), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            unidad_path    = tmp / safe_name(unidad.get("name", "unidad.docx"))
            plantilla_path = tmp / safe_name(plantilla.get("name", "plantilla.docx"))
            salida_path    = tmp / safe_name(nombre_salida)

            write_b64(unidad_path, unidad["base64"])
            write_b64(plantilla_path, plantilla["base64"])

            inter_path = None
            if interacc and interacc.get("base64"):
                inter_path = tmp / safe_name(interacc.get("name", "interacciones.docx"))
                write_b64(inter_path, interacc["base64"])

            stdout, stderr = _ejecutar_conversor(
                unidad_path, plantilla_path, inter_path, salida_path
            )

            if not salida_path.exists():
                return jsonify({
                    "ok": False,
                    "error": "El conversor no generó el DOCX de salida.",
                    "stdout": stdout, "stderr": stderr
                }), 500

            out_b64 = base64.b64encode(salida_path.read_bytes()).decode()

        return jsonify({
            "ok": True,
            "nombre": salida_path.name,
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "base64": out_b64,
            "log": ["Conversión completada.", "STDOUT:", stdout, "STDERR:", stderr]
        })

    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "La conversión superó el tiempo máximo de 300 s."}), 504
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()}), 500


# ── /convertir_lote  (varias unidades → un único DOCX) ───────────────────────

@app.route("/convertir_lote", methods=["POST"])
def convertir_lote():
    log = []
    try:
        data = request.get_json(force=True)

        nombre_salida   = data.get("nombreSalida") or "curso_papel.docx"
        archivos        = data.get("archivos", {})
        plantilla       = archivos.get("plantilla")
        unidades_lista  = archivos.get("unidades") or []

        if not plantilla or not plantilla.get("base64"):
            return jsonify({"ok": False, "error": "Falta la plantilla maquetada."}), 400
        if not unidades_lista:
            return jsonify({"ok": False, "error": "No se han enviado unidades."}), 400

        rutas_salida = []

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            plantilla_path = tmp / safe_name(plantilla.get("name", "plantilla.docx"))
            write_b64(plantilla_path, plantilla["base64"])

            for i, u in enumerate(unidades_lista):
                unidad   = u.get("unidad")
                interacc = u.get("interacciones")
                numero   = u.get("numero", i + 1)

                if not unidad or not unidad.get("base64"):
                    return jsonify({"ok": False, "error": f"Falta el archivo de la Unidad {numero}."}), 400

                log.append(f"Convirtiendo Unidad {numero}…")

                unit_dir = tmp / f"u{i}"
                unit_dir.mkdir()

                unidad_path = unit_dir / safe_name(unidad.get("name", f"unidad_{numero}.docx"))
                salida_path = unit_dir / f"salida_{i}.docx"
                write_b64(unidad_path, unidad["base64"])

                inter_path = None
                if interacc and interacc.get("base64"):
                    inter_path = unit_dir / safe_name(interacc.get("name", f"inter_{numero}.docx"))
                    write_b64(inter_path, interacc["base64"])

                stdout, stderr = _ejecutar_conversor(
                    unidad_path, plantilla_path, inter_path, salida_path
                )

                if not salida_path.exists():
                    return jsonify({
                        "ok": False,
                        "error": f"El conversor falló en la Unidad {numero}.",
                        "stdout": stdout, "stderr": stderr
                    }), 500

                # Copiamos el resultado a la raíz del tmpdir para que no desaparezca
                destino = tmp / f"salida_{i}.docx"
                destino.write_bytes(salida_path.read_bytes())
                rutas_salida.append(destino)
                log.append(f"Unidad {numero} completada.")

            log.append("Fusionando documentos…")
            docx_bytes = _fusionar_docx(rutas_salida)

        nombre_salida = safe_name(nombre_salida)
        if not nombre_salida.lower().endswith(".docx"):
            nombre_salida += ".docx"

        log.append(f"Listo: {nombre_salida} ({len(unidades_lista)} unidades)")

        return jsonify({
            "ok": True,
            "nombre": nombre_salida,
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "base64": base64.b64encode(docx_bytes).decode(),
            "log": log
        })

    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "La conversión superó el tiempo máximo de 300 s.", "log": log}), 504
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc(), "log": log}), 500


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ejecutar_conversor(unidad_path, plantilla_path, inter_path, salida_path):
    """
    Llama al script conversor vía subprocess.
    Firma del script:
      3 args → unidad ejemplo salida
      4 args → unidad ejemplo interacciones salida
    """
    if inter_path and inter_path.exists():
        cmd = [sys.executable, str(CONVERSOR),
               str(unidad_path), str(plantilla_path),
               str(inter_path), str(salida_path)]
    else:
        cmd = [sys.executable, str(CONVERSOR),
               str(unidad_path), str(plantilla_path),
               str(salida_path)]

    result = subprocess.run(
        cmd,
        cwd=str(Path(__file__).parent),
        capture_output=True,
        text=True,
        timeout=300
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"El script falló (código {result.returncode}).\n"
            f"STDERR: {result.stderr[:1000]}"
        )

    return result.stdout, result.stderr


def _fusionar_docx(rutas: list) -> bytes:
    """
    Une varios DOCX en uno insertando un salto de página entre ellos.
    Trabaja directamente con el ZIP/XML para no necesitar dependencias extra.
    """
    if len(rutas) == 1:
        return rutas[0].read_bytes()

    with zipfile.ZipFile(rutas[0], "r") as z:
        base_files = {name: z.read(name) for name in z.namelist()}

    doc_xml = base_files["word/document.xml"].decode("utf-8")

    def _body_content(xml):
        m = re.search(r"<w:body>([\s\S]*)</w:body>", xml)
        if not m:
            return ""
        content = m.group(1)
        # Quitamos la sectPr final (define el layout de página; la del primer doc manda)
        return re.sub(r"<w:sectPr[\s\S]*?</w:sectPr>\s*$", "", content).rstrip()

    def _sectpr(xml):
        m = re.search(r"(<w:sectPr[\s\S]*?</w:sectPr>)\s*</w:body>", xml)
        return m.group(1) if m else ""

    PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    body = _body_content(doc_xml)

    for ruta in rutas[1:]:
        with zipfile.ZipFile(ruta, "r") as z:
            other_xml = z.read("word/document.xml").decode("utf-8")
        body += PAGE_BREAK + _body_content(other_xml)

    new_doc_xml = re.sub(
        r"<w:body>[\s\S]*</w:body>",
        f"<w:body>{body}{_sectpr(doc_xml)}</w:body>",
        doc_xml
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in base_files.items():
            if name == "word/document.xml":
                zout.writestr(name, new_doc_xml.encode("utf-8"))
            else:
                zout.writestr(name, data)

    return buf.getvalue()


def write_b64(path: Path, b64: str):
    path.write_bytes(base64.b64decode(b64))


def safe_name(name: str) -> str:
    for ch in '\\/:*?"<>|#{}%~&':
        name = name.replace(ch, "_")
    return name.strip() or "archivo"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
