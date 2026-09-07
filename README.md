# DocuMation — microservicio de recibos para Render.com

Este microservicio reemplaza el renderizado aproximado de PDF por **conversión real con LibreOffice Calc**, ejecutándose en un contenedor Docker desplegable en Render.

## Algoritmo

Para cada hoja del Excel:

- ignora `Modelo`, `SICOSS`, `Resumen`, `CUSS`, `Hoja6`, `SAC_VAC`;
- ignora hojas cuyo nombre sea únicamente numérico;
- genera `B80:G153` como `Original`;
- genera `B2:G77` como `Duplicado`;
- configura el área de impresión del rango;
- fuerza una página de ancho y una de alto mediante las propiedades de impresión de LibreOffice;
- exporta a PDF mediante `calc_pdf_Export`.

Esto utiliza el motor real de Calc, no una reconstrucción de celdas con `pdf-lib`.

## Endpoints

### `POST /api/payslips/upload`

Compatible con el endpoint utilizado por el algoritmo original.

Multipart:

- `file`: Excel `.xlsx`/`.xls`
- `month`: `YYYY-MM`

### `POST /api/generar-recibo`

Alias para conservar compatibilidad con la primera versión del microservicio.

### `GET /health`

Health check para Render.

## Almacenamiento

No existe almacenamiento persistente. El Excel se guarda únicamente en un directorio temporal durante la petición porque LibreOffice necesita una ruta de archivo; al terminar la petición, el directorio temporal se elimina. Los PDFs y el ZIP tampoco se persisten.

Si el sistema necesita guardar los recibos a largo plazo, esa responsabilidad debe estar en Supabase Storage, S3, etc., no en el filesystem del contenedor de Render.

## Despliegue en Render

1. Subir este directorio a un repositorio Git.
2. Crear un Web Service en Render desde ese repositorio.
3. Seleccionar Docker.
4. Render usará `Dockerfile` y expondrá el puerto mediante `PORT`.
5. Health check: `/health`.

## Prueba local

```bash
docker build -t documation-payslips .
docker run --rm -p 10000:10000 documation-payslips
```

Luego:

```bash
curl -X POST http://localhost:10000/api/payslips/upload \
  -F "file=@Recibos Sueldos -para prueba.xls.xlsx" \
  -F "month=2026-06" \
  -o recibos_2026-06.zip
```
