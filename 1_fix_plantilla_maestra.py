"""
Corrección ÚNICA sobre la plantilla maestra de recibos.
Arregla los dos gráficos de torta (Duplicado fila 62 y Original fila 138)
y elimina las referencias a libros externos congelados.
"""

import openpyxl

ENTRADA = "plantilla_original.xlsx"
SALIDA = "plantilla_maestra_fixed.xlsx"


def fix_hoja(ws):
    nombre_hoja = ws.title
    sheet_ref = f"'{nombre_hoja}'"

    # 1. Configuración de página base
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    if ws.sheet_properties.pageSetUpPr is None:
        from openpyxl.worksheet.properties import PageSetupProperties
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25

    # 2. Re-vincular gráficos locales:
    # Chart 0: Duplicado (anclado en fila ~62) -> datos I65:J70
    # Chart 1: Original  (anclado en fila ~138) -> datos I141:J146
    for chart in ws._charts:
        row_anchor = chart.anchor._from.row if hasattr(chart.anchor, "_from") else 0
        is_original = row_anchor > 100

        rango_cat = "$I$141:$I$146" if is_original else "$I$65:$I$70"
        rango_val = "$J$141:$J$146" if is_original else "$J$65:$J$70"

        for serie in chart.series:
            if serie.cat and serie.cat.strRef:
                serie.cat.strRef.f = f"{sheet_ref}!{rango_cat}"
                serie.cat.strRef.strCache = None
            if serie.val and serie.val.numRef:
                serie.val.numRef.f = f"{sheet_ref}!{rango_val}"
                serie.val.numRef.numCache = None


def main():
    wb = openpyxl.load_workbook(ENTRADA)
    for nombre_hoja in wb.sheetnames:
        fix_hoja(wb[nombre_hoja])
    wb.save(SALIDA)
    print(f"Plantilla corregida guardada exitosamente en: {SALIDA}")


if __name__ == "__main__":
    main()