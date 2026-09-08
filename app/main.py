import asyncio
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

EXCLUDED_SHEETS = {"Modelo", "SICOSS", "Resumen", "CUSS", "Hoja6", "SAC_VAC"}
ORIGINAL_RANGE = "B80:G153"
DUPLICATE_RANGE = "B2:G77"
MAX_FILE_MB = int(os.getenv("MAX_FILE_MB", "50"))

app = FastAPI(title="DocuMation Recibos PDF", version="2.0.0")


def sanitize_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", " ", value)
    return (value[:150] or "recibo")


def valid_sheet(name: str) -> bool:
    return name not in EXCLUDED_SHEETS and not re.fullmatch(r"\d+", name or "")


def validate_month(month: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}", month or ""))


def start_libreoffice(profile_dir: Path, port: int = 2002):
    profile_uri = profile_dir.resolve().as_uri()
    cmd = [
        "soffice",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation={profile_uri}",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    return subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def wait_for_uno(port: int, timeout: float = 20):
    # UNO is imported from the python3-uno Debian package installed in the image.
    import uno
    from com.sun.star.connection import NoConnectException

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
            )
            return ctx
        except NoConnectException as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"No se pudo conectar con LibreOffice: {last_error}")


def uno_path(path: Path) -> str:
    import uno
    return uno.systemPathToFileUrl(str(path.resolve()))


def export_sheet_range(ctx, input_path: Path, output_path: Path, sheet_name: str, range_a1: str):
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    props = [
        make_prop("Hidden", True),
        make_prop("ReadOnly", True),
        make_prop("UpdateDocMode", 3),
    ]
    doc = desktop.loadComponentFromURL(uno_path(input_path), "_blank", 0, tuple(props))
    if doc is None:
        raise RuntimeError("LibreOffice no pudo abrir el archivo Excel.")

    try:
        sheets = doc.getSheets()
        names = list(sheets.getElementNames())
        if sheet_name not in names:
            raise RuntimeError(f"No existe la hoja '{sheet_name}'.")

        target = sheets.getByName(sheet_name)

        # Solo dejamos visible la hoja objetivo para que el PDF contenga un recibo.
        for name in names:
            sheets.getByName(name).IsVisible = (name == sheet_name)

        target.setPrintAreas((target.getCellRangeByName(range_a1).getRangeAddress(),))
        doc.getCurrentController().setActiveSheet(target)

        # Equivalente funcional a Excel: FitToPagesWide=1 / FitToPagesTall=1.
        page_style_name = target.getPropertyValue("PageStyle")
        page_styles = doc.getStyleFamilies().getByName("PageStyles")
        page_style = page_styles.getByName(page_style_name)
        if page_style.getPropertySetInfo().hasPropertyByName("ScaleToPagesX"):
            page_style.setPropertyValue("ScaleToPagesX", 1)
        if page_style.getPropertySetInfo().hasPropertyByName("ScaleToPagesY"):
            page_style.setPropertyValue("ScaleToPagesY", 1)
        if page_style.getPropertySetInfo().hasPropertyByName("PageScale"):
            page_style.setPropertyValue("PageScale", 100)

        filter_props = (
            make_prop("FilterName", "calc_pdf_Export"),
            make_prop("Overwrite", True),
        )
        doc.storeToURL(uno_path(output_path), filter_props)
    finally:
        doc.close(True)


def make_prop(name, value):
    import uno
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def render_workbook(input_path: Path, work_dir: Path, month: str):
    profile = work_dir / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    port = 2002
    proc = start_libreoffice(profile, port)
    try:
        ctx = wait_for_uno(port)
        # Obtain sheet names through LibreOffice itself, preserving the workbook's real structure.
        import uno
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno_path(input_path), "_blank", 0, (make_prop("Hidden", True), make_prop("ReadOnly", True)))
        if doc is None:
            raise RuntimeError("No se pudo abrir el XLSX con LibreOffice.")
        try:
            sheet_names = [n for n in doc.getSheets().getElementNames() if valid_sheet(n)]
        finally:
            doc.close(True)

        generated = []
        for sheet_name in sheet_names:
            safe = sanitize_filename(sheet_name)
            for kind, rng in (("Original", ORIGINAL_RANGE), ("Duplicado", DUPLICATE_RANGE)):
                pdf_path = work_dir / f"{safe} - {kind}.pdf"
                export_sheet_range(ctx, input_path, pdf_path, sheet_name, rng)
                generated.append({
                    "sheet": sheet_name,
                    "type": kind,
                    "range": rng,
                    "file": pdf_path.name,
                    "path": pdf_path,
                })
        return generated
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


async def process(file: UploadFile, month: str):
    if not validate_month(month):
        raise HTTPException(400, "El campo 'month' es obligatorio y debe tener formato YYYY-MM.")
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "El campo 'file' debe ser un Excel .xlsx o .xls.")

    data = await file.read()
    if not data:
        raise HTTPException(400, "El archivo Excel está vacío.")
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        raise HTTPException(413, f"El Excel supera el límite de {MAX_FILE_MB} MB.")

    # Solo almacenamiento transitorio durante la petición. No se persiste nada.
    with tempfile.TemporaryDirectory(prefix="documation-") as td:
        work_dir = Path(td)
        input_ext = ".xlsx" if file.filename.lower().endswith(".xlsx") else ".xls"
        input_path = work_dir / f"source{input_ext}"
        input_path.write_bytes(data)

        try:
            generated = await asyncio.to_thread(render_workbook, input_path, work_dir, month)
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(500, f"Error convirtiendo el Excel con LibreOffice: {exc}") from exc

        if not generated:
            raise HTTPException(400, "El Excel no contiene hojas válidas para generar recibos.")

        manifest = {
            "month": month,
            "originalRange": ORIGINAL_RANGE,
            "duplicateRange": DUPLICATE_RANGE,
            "renderer": "LibreOffice Calc",
            "generated": [
                {k: v for k, v in item.items() if k != "path"}
                for item in generated
            ],
        }

        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, "w", ZIP_DEFLATED) as zf:
            for item in generated:
                zf.write(item["path"], arcname=item["file"])
            zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        zip_buffer.seek(0)

        base = sanitize_filename(re.sub(r"\.xlsx?$", "", file.filename, flags=re.I))
        download_name = f"{base}_{month}_recibos.zip"
        headers = {
            "Content-Disposition": f'attachment; filename="{download_name}"',
            "X-Generated-Count": str(len(generated)),
        }
        return Response(zip_buffer.getvalue(), media_type="application/zip", headers=headers)


@app.get("/health")
def health():
    return {"status": "ok", "renderer": "LibreOffice Calc"}


@app.post("/api/payslips/upload")
async def upload_payslips(file: UploadFile = File(...), month: str = Form(...)):
    return await process(file, month)


@app.post("/api/generar-recibo")
async def generar_recibo(file: UploadFile = File(...), month: str = Form(...)):
    # Alias compatible con la primera versión del microservicio.
    return await process(file, month)
