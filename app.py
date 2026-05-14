import base64
import os
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path

from flask import Flask, request, jsonify

app = Flask(__name__)


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "ok": True,
        "service": "conversor-papel",
        "status": "online"
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True})


@app.route("/convertir", methods=["POST"])
def convertir():
    try:
        data = request.get_json(force=True)

        nombre_salida = data.get("nombreSalida") or "unidad_papel.docx"
        archivos = data.get("archivos", {})

        unidad = archivos.get("unidad")
        plantilla = archivos.get("plantilla")
        interacciones = archivos.get("interacciones")

        if not unidad:
            return jsonify({
                "ok": False,
                "error": "Falta archivo de unidad."
            }), 400

        if not plantilla:
            return jsonify({
                "ok": False,
                "error": "Falta archivo de plantilla."
            }), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            unidad_path = tmp / safe_name(unidad.get("name", "unidad.docx"))
            plantilla_path = tmp / safe_name(plantilla.get("name", "plantilla.docx"))
            salida_path = tmp / safe_name(nombre_salida)

            write_b64(unidad_path, unidad["base64"])
            write_b64(plantilla_path, plantilla["base64"])

            # Tu script original funciona así:
            # python conversor_papel.py UNIDAD.pdf EJEMPLO_MAQUETADO.docx [SALIDA.docx]
            # Por eso lo ejecutamos por consola para respetar su comportamiento.
            cmd = [
                sys.executable,
                str(Path(__file__).parent / "conversor_papel.py"),
                str(unidad_path),
                str(plantilla_path),
                str(salida_path)
            ]

            # Si más adelante adaptamos el Python para recibir interacciones,
            # lo añadiremos aquí. De momento no lo pasamos para no romper la firma original.

            result = subprocess.run(
                cmd,
                cwd=str(Path(__file__).parent),
                capture_output=True,
                text=True,
                timeout=300
            )

            if result.returncode != 0:
                return jsonify({
                    "ok": False,
                    "error": "El script Python falló.",
                    "stdout": result.stdout,
                    "stderr": result.stderr
                }), 500

            if not salida_path.exists():
                return jsonify({
                    "ok": False,
                    "error": "El conversor Python no generó el DOCX de salida.",
                    "stdout": result.stdout,
                    "stderr": result.stderr
                }), 500

            out_b64 = base64.b64encode(salida_path.read_bytes()).decode("utf-8")

            return jsonify({
                "ok": True,
                "nombre": salida_path.name,
                "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "base64": out_b64,
                "log": [
                    "Conversión realizada con motor Python.",
                    "Unidad: " + unidad_path.name,
                    "Plantilla: " + plantilla_path.name,
                    "Salida: " + salida_path.name,
                    "STDOUT:",
                    result.stdout,
                    "STDERR:",
                    result.stderr
                ]
            })

    except subprocess.TimeoutExpired:
        return jsonify({
            "ok": False,
            "error": "La conversión superó el tiempo máximo de 300 segundos."
        }), 504

    except Exception as e:
        return jsonify({
            "ok": False,
            "error": str(e),
            "trace": traceback.format_exc()
        }), 500


def write_b64(path: Path, b64: str):
    path.write_bytes(base64.b64decode(b64))


def safe_name(name: str) -> str:
    bad = '\\/:*?"<>|#{}%~&'
    for ch in bad:
        name = name.replace(ch, "_")
    return name.strip() or "archivo"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
