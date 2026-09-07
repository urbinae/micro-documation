"""
Microservicio de generación de recibos en PDF.
Escribe automáticamente en Duplicado (filas 2-77) y Original (filas 80-153)
y retorna ambos PDFs o el que se solicite por parámetro.
"""

import os
import subprocess
import tempfile
import uuid

import openpyxl
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)

PLANTILLA_MAESTRA = os.path.join(os.path.dirname(__file__), "plantilla_maestra_fixed.xlsx")
OFFSET_ORIGINAL = 76  # Desplazamiento exacto de filas entre Duplicado y Original

# Mapeo de celdas para el Duplicado (filas 2 a 77)
CELDAS_BASE = {
    "nombre_apellido": "C8",
    "periodo_abonado": "B8",
    "cuil": "G7",
    "obra_social": "G9",
    "banco": "B11",
    "periodo_seg_soc": "C11",
    "fecha_deposito": "D11",
    "tarea": "E11",
    "fecha_ingreso": "F11",
    "rem_basica": "G11",
    "remuneracion": "F13",
    "a_cuenta_futuros_aumentos": "F18",
    "importe_jubilacion": "G22",
    "importe_inssjp": "G23",
    "importe_obra_social_desc": "G24",
}

# Tabla de gráfico Duplicado
CELDAS_GRAFICO_BASE = {
    "sueldo_neto": "J65",
    "seguridad_social": "J66",
    "obra_social_total": "J67",
    "art": "J68",
    "costo_sindical": "J69",
    "scvo": "J70",
}


def desplazar_celda(celda: str, offset: int) -> str:
    col = "".join([c for c in celda if c.isalpha()])
    fila = int("".join([c for c in celda if c.isdigit()]))
    return f"{col}{fila + offset}"


def rellenar_datos(ws, datos: dict):
    # 1. Llenar Duplicado
    for campo, celda in CELDAS_BASE.items():
        if campo in datos:
            ws[celda] = datos[campo]
    for campo, celda in CELDAS_GRAFICO_BASE.items():
        if campo in datos:
            ws[celda] = datos[campo]

    # 2. Llenar Original con offset de 76 filas
    for campo, celda in CELDAS_BASE.items():
        if campo in datos:
            ws[desplazar_celda(celda, OFFSET_ORIGINAL)] = datos[campo]
    for campo, celda in CELDAS_GRAFICO_BASE.items():
        if campo in datos:
            ws[desplazar_celda(celda, OFFSET_ORIGINAL)] = datos[campo]


def exportar_rango_pdf(wb, ws, print_area: str, out_pdf_path: str, tmp_dir: str):
    ws.print_area = print_area
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    temp_xlsx = os.path.join(tmp_dir, f"export_{uuid.uuid4().hex}.xlsx")
    wb.save(temp_xlsx)

    resultado = subprocess.run(
        [
            "soffice", "--headless", "--norestore",
            "--convert-to", "pdf", "--outdir", tmp_dir, temp_xlsx,
        ],
        capture_output=True, text=True, timeout=60,
    )
    if resultado.returncode != 0:
        raise RuntimeError(f"Fallo la conversión de {print_area} a PDF: {resultado.stderr}")

    generated_pdf = temp_xlsx.replace(".xlsx", ".pdf")
    os.replace(generated_pdf, out_pdf_path)


@app.post("/generar-recibo")
def generar_recibo():
    datos = request.get_json(force=True)
    if not datos:
        return jsonify({"error": "Body JSON vacío"}), 400

    tipo = request.args.get("tipo", "duplicado").lower()  # "duplicado" o "original"

    with tempfile.TemporaryDirectory() as tmp:
        wb = openpyxl.load_workbook(PLANTILLA_MAESTRA)
        ws = wb.active

        rellenar_datos(ws, datos)

        rango = "B2:G77" if tipo == "duplicado" else "B80:G153"
        pdf_name = f"recibo_{tipo}_{uuid.uuid4().hex}.pdf"
        pdf_tmp_path = os.path.join(tmp, pdf_name)

        try:
            exportar_rango_pdf(wb, ws, rango, pdf_tmp_path, tmp)
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 500

        pdf_final = os.path.join(tempfile.gettempdir(), pdf_name)
        os.replace(pdf_tmp_path, pdf_final)

    return send_file(pdf_final, mimetype="application/pdf",
                      as_attachment=True, download_name=f"recibo_{tipo}.pdf")


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))