"""
Microservicio de generación de recibos PDF.

Reemplaza al script de PowerShell (Excel COM Object) usando LibreOffice
vía la API UNO. La diferencia clave frente a `exceljs`/`pdf-lib` es que
acá NUNCA se reconstruye el archivo: LibreOffice abre el .xlsx original
tal cual (con sus gráficos de torta e imágenes nativos) y exporta
directamente el rango pedido a PDF, igual que hacía
`Range.ExportAsFixedFormat` en Excel COM.

No se persiste nada en disco de forma duradera: los archivos temporales
que exige la API de LibreOffice se leen a memoria y se borran
inmediatamente después de cada exportación.
"""

import os
import re
import tempfile
import uuid

import requests
import uno
from com.sun.star.beans import PropertyValue
from flask import Flask, jsonify, request

app = Flask(__name__)

# --- Misma lógica de exclusión que el script de PowerShell -----------------
EXCLUDED_NAMES = {"Modelo", "SICOSS", "Resumen", "CUSS", "Hoja6", "SAC_VAC"}
NUMERIC_ONLY_RE = re.compile(r"^\d+$")

# --- Mismos rangos que ExportAsFixedFormat en el script original ----------
RANGES = {
    "Original": "B80:G153",
    "Duplicado": "B2:G77",
}

UNO_HOST = os.environ.get("UNO_HOST", "localhost")
UNO_PORT = os.environ.get("UNO_PORT", "2002")

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
SUPABASE_BUCKET = os.environ.get("SUPABASE_BUCKET", "recibos")


def is_valid_sheet(name: str) -> bool:
    if name in EXCLUDED_NAMES:
        return False
    if NUMERIC_ONLY_RE.match(name):
        return False
    return True


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect_to_soffice():
    """Conecta al proceso de LibreOffice que corre en background (start.sh)."""
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve(
        f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext"
    )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop


def open_workbook(desktop, path: str):
    url = uno.systemPathToFileUrl(path)
    props = (make_prop("Hidden", True),)
    return desktop.loadComponentFromURL(url, "_blank", 0, props)


def fit_sheet_to_one_page(doc, sheet):
    """Equivalente a PageSetup.FitToPagesWide/Tall = 1 del script original."""
    try:
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        style = page_styles.getByName(sheet.PageStyle)
        style.ScaleToPagesX = 1
        style.ScaleToPagesY = 1
    except Exception:
        # Si el estilo de página no soporta escalado, seguimos igual:
        # la selección exportada no depende estrictamente de esto.
        pass


def export_range_to_pdf(doc, sheet, range_address: str, out_path: str):
    """Exporta SOLO el rango indicado a PDF, preservando gráficos e imágenes
    porque LibreOffice renderiza el documento real, no una reconstrucción."""
    cell_range = sheet.getCellRangeByName(range_address)
    doc.CurrentController.setActiveSheet(sheet)
    doc.CurrentController.select(cell_range)

    filter_data = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        (make_prop("Selection", cell_range),),
    )
    export_props = (
        make_prop("FilterName", "calc_pdf_Export"),
        make_prop("FilterData", filter_data),
    )
    url = uno.systemPathToFileUrl(out_path)
    doc.storeToURL(url, export_props)


def upload_pdf_to_supabase(pdf_bytes: bytes, filename: str, month: str) -> str:
    path = f"{month}/{filename}"
    resp = requests.post(
        f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_BUCKET}/{path}",
        headers={
            "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
            "Content-Type": "application/pdf",
            "x-upsert": "true",
        },
        data=pdf_bytes,
        timeout=30,
    )
    resp.raise_for_status()
    return path


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generate-receipts", methods=["POST"])
def generate_receipts():
    if "file" not in request.files:
        return jsonify({"error": "falta el archivo 'file'"}), 400
    month = request.form.get("month")
    if not month:
        return jsonify({"error": "falta el campo 'month'"}), 400

    uploaded = []
    skipped = []

    # Directorio temporal: se borra completo al salir del `with`, nunca
    # queda nada persistido en disco.
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, f"{uuid.uuid4()}.xlsx")
        request.files["file"].save(input_path)

        desktop = connect_to_soffice()
        doc = open_workbook(desktop, input_path)
        try:
            sheets = doc.Sheets
            for i in range(sheets.Count):
                sheet = sheets.getByIndex(i)
                name = sheet.Name

                if not is_valid_sheet(name):
                    skipped.append(name)
                    continue

                fit_sheet_to_one_page(doc, sheet)

                for label, range_address in RANGES.items():
                    out_path = os.path.join(tmpdir, f"{uuid.uuid4()}.pdf")
                    export_range_to_pdf(doc, sheet, range_address, out_path)

                    with open(out_path, "rb") as f:
                        pdf_bytes = f.read()
                    os.remove(out_path)  # nunca queda en disco

                    filename = f"{name} - {label}.pdf"
                    storage_path = upload_pdf_to_supabase(pdf_bytes, filename, month)
                    uploaded.append(storage_path)
        finally:
            doc.close(False)
        os.remove(input_path)

    return jsonify({"generated": uploaded, "skipped_sheets": skipped})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)