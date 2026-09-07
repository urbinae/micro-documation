"""
Corrección ÚNICA sobre la plantilla maestra de recibos.
Correr UNA sola vez sobre el archivo original y guardar el resultado
como la nueva plantilla que usará el microservicio en cada request.

Qué arregla:
1. El gráfico de torta ("Costo Total empleador") está vinculado a un
   workbook externo que no existe (referencia tipo '[1]Latsague V'!$I$65...).
   Eso congela el gráfico con el último valor guardado, sin importar qué
   celdas cambies después. Lo re-vinculamos a las celdas locales I65:J70
   de la propia hoja, que ya tienen los mismos datos en espejo.
2. No hay área de impresión / ajuste de página definidos, lo que hace que
   el PDF se pagine mal. Se define un área de impresión razonable y
   ajuste a ancho de página.

IMPORTANTE: los rangos de celdas (I65:J70, A1:K151, etc.) están hardcodeados
según la plantilla real inspeccionada. Si tu plantilla tiene una segunda
copia del recibo (duplicado) más abajo en la hoja, repetí el mismo
re-vinculado de gráfico para esa segunda copia con su propio rango de celdas
(hay que verificar sus coordenadas exactas — no vinieron incluidas en esta
inspección).
"""

import openpyxl
import sys

ENTRADA = "plantilla_original.xlsx"
SALIDA = "plantilla_maestra_fixed.xlsx"


def fix_hoja(ws):
    nombre_hoja = ws.title
    sheet_ref = f"'{nombre_hoja}'"

    # --- 1. Área de impresión y ajuste de página ---
    # Ajustá A1:K151 al rango real de tu plantilla si difiere.
    ws.print_area = "A1:K151"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 0 = alto libre, ajustado solo por ancho
    if ws.sheet_properties.pageSetUpPr is None:
        from openpyxl.worksheet.properties import PageSetupProperties
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3

    # --- 2. Re-vincular el/los gráfico(s) de torta a celdas locales ---
    for chart in ws._charts:
        for serie in chart.series:
            # Categorías (nombres: "Sueldo neto", "Seguridad Social", etc.)
            if serie.cat and serie.cat.strRef:
                serie.cat.strRef.f = f"{sheet_ref}!$I$65:$I$70"
                serie.cat.strRef.strCache = None  # se recalcula al abrir/convertir
            # Valores
            if serie.val and serie.val.numRef:
                serie.val.numRef.f = f"{sheet_ref}!$J$65:$J$70"
                serie.val.numRef.numCache = None


def main():
    wb = openpyxl.load_workbook(ENTRADA)
    for nombre_hoja in wb.sheetnames:
        fix_hoja(wb[nombre_hoja])
    wb.save(SALIDA)
    print(f"Plantilla corregida guardada en: {SALIDA}")


if __name__ == "__main__":
    main()
