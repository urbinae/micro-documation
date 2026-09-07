// /api/generar-recibo.js (Vercel Serverless Function)
import { createClient } from "@supabase/supabase-js";

const GENERADOR_URL = process.env.RECIBOS_SERVICE_URL; // ej: https://recibos.fly.dev
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

async function obtenerPdfDeMicroservicio(datosRecibo, tipo) {
  const resp = await fetch(`${GENERADOR_URL}/generar-recibo?tipo=${tipo}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(datosRecibo),
  });

  if (!resp.ok) {
    const errorText = await resp.text();
    throw new Error(`Fallo generador (${tipo}): ${errorText}`);
  }

  return Buffer.from(await resp.arrayBuffer());
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const datos = req.body;
  if (!datos || !datos.cuil) {
    return res.status(400).json({ error: "Datos del recibo incompletos (CUIL requerido)" });
  }

  const periodo = datos.periodo_abonado ? datos.periodo_abonado.replace(/[\/\s:]/g, "-") : "periodo";

  try {
    // Generar ambos PDFs en paralelo (Duplicado B2:G77 y Original B80:G153)
    const [duplicadoBuffer, originalBuffer] = await Promise.all([
      obtenerPdfDeMicroservicio(datos, "duplicado"),
      obtenerPdfDeMicroservicio(datos, "original"),
    ]);

    const pathDuplicado = `duplicados/${datos.cuil}_${periodo}_duplicado.pdf`;
    const pathOriginal = `originals/${datos.cuil}_${periodo}_original.pdf`;

    // Subida concurrente a Supabase Storage
    const [upDuplicado, upOriginal] = await Promise.all([
      supabase.storage.from("payslips").upload(pathDuplicado, duplicadoBuffer, {
        contentType: "application/pdf",
        upsert: true,
      }),
      supabase.storage.from("payslips").upload(pathOriginal, originalBuffer, {
        contentType: "application/pdf",
        upsert: true,
      }),
    ]);

    if (upDuplicado.error) throw new Error(upDuplicado.error.message);
    if (upOriginal.error) throw new Error(upOriginal.error.message);

    const { data: urlDuplicado } = supabase.storage.from("payslips").getPublicUrl(pathDuplicado);
    const { data: urlOriginal } = supabase.storage.from("payslips").getPublicUrl(pathOriginal);

    return res.status(200).json({
      success: true,
      cuil: datos.cuil,
      periodo: periodo,
      duplicadoUrl: urlDuplicado.publicUrl,
      originalUrl: urlOriginal.publicUrl,
    });
  } catch (err) {
    console.error("Error en api/generar-recibo:", err);
    return res.status(500).json({ error: "Error generando recibos", detalle: err.message });
  }
}