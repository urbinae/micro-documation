"""
Microservicio de generación de recibos en PDF.

Recibe los datos de un recibo por POST, los escribe sobre la plantilla
maestra ya corregida (ver 1_fix_plantilla_maestra.py), y devuelve el PDF
resultante generado con LibreOffice headless (preserva imagen y gráfico
de torta tal cual están en el Excel).

Pensado para correr en un contenedor pequeño (Fly.io, Render, etc.) con
Python + LibreOffice instalados — NO corre dentro de Vercel, que no
soporta LibreOffice. Tu función de Vercel llama a este servicio por HTTP.

Dependencias: flask, openpyxl, gunicorn (ver requirements.txt)
"""

import os
import subprocess
import tempfile
import uuid

import openpyxl
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)

PLANTILLA_MAESTRA = os.path.join(os.path.dirname(__file__), "plantilla_maestra_fixed.xlsx")

# Nombre de la hoja a usar como base dentro de la plantilla maestra.
# Si tenés una hoja distinta por empleado en la plantilla original, en
# producción probablemente quieras UNA sola hoja "molde" y clonarla,
# en vez de tener una hoja por persona como en el archivo de prueba.
HOJA_MOLDE = "Molde"

# Mapeo de campos del recibo -> celda. Verificado contra la plantilla real.
CELDAS = {
    "nombre_apellido": "C8",
    "periodo_abonado": "B8",       # fecha
    "cuil": "G7",
    "obra_social": "G9",
    "banco": "B11",
    "periodo_seg_soc": "C11",      # fecha
    "fecha_deposito": "D11",       # fecha
    "tarea": "E11",
    "fecha_ingreso": "F11",        # fecha
    "rem_basica": "G11",
    "remuneracion": "F13",
    "a_cuenta_futuros_aumentos": "F18",
    "importe_jubilacion": "G22",
    "importe_inssjp": "G23",
    "importe_obra_social_desc": "G24",
}

# Tabla que alimenta el gráfico de torta (celdas I65:J70, ya re-vinculadas
# localmente por el script de arreglo de la plantilla).
CELDAS_GRAFICO = {
    "sueldo_neto": "J65",
    "seguridad_social": "J66",
    "obra_social_total": "J67",
    "art": "J68",
    "costo_sindical": "J69",
    "scvo": "J70",
}

# TODO: si tu plantilla tiene una segunda copia del recibo (duplicado) más
# abajo en la misma hoja, agregá acá el mismo mapeo con el offset de filas
# correspondiente y escribilo también en `generar_pdf()`.


def rellenar_datos(ws, datos: dict):
    for campo, celda in CELDAS.items():
        if campo in datos:
            ws[celda] = datos[campo]
    for campo, celda in CELDAS_GRAFICO.items():
        if campo in datos:
            ws[celda] = datos[campo]


def convertir_a_pdf(xlsx_path: str, out_dir: str) -> str:
    resultado = subprocess.run(
        [
            "soffice", "--headless", "--norestore",
            "--convert-to", "pdf", "--outdir", out_dir, xlsx_path,
        ],
        capture_output=True, text=True, timeout=60,
    )
    if resultado.returncode != 0:
        raise RuntimeError(f"Fallo la conversión a PDF: {resultado.stderr}")
    return os.path.join(out_dir, os.path.basename(xlsx_path).replace(".xlsx", ".pdf"))


@app.post("/generar-recibo")
def generar_recibo():
    datos = request.get_json(force=True)
    if not datos:
        return jsonify({"error": "Body JSON vacío"}), 400

    with tempfile.TemporaryDirectory() as tmp:
        wb = openpyxl.load_workbook(PLANTILLA_MAESTRA)
        ws = wb[HOJA_MOLDE] if HOJA_MOLDE in wb.sheetnames else wb[wb.sheetnames[0]]
        rellenar_datos(ws, datos)

        xlsx_path = os.path.join(tmp, f"recibo_{uuid.uuid4().hex}.xlsx")
        wb.save(xlsx_path)

        try:
            pdf_path = convertir_a_pdf(xlsx_path, tmp)
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 500

        # send_file necesita que el archivo siga existiendo al momento de
        # enviarlo; lo copiamos fuera del TemporaryDirectory antes de que
        # se borre al salir del "with".
        pdf_final = pdf_path.replace(tmp, tempfile.gettempdir())
        os.replace(pdf_path, pdf_final)

    return send_file(pdf_final, mimetype="application/pdf",
                      as_attachment=True, download_name="recibo.pdf")


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
