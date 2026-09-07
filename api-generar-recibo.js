// /api/generar-recibo.js  (Vercel serverless function - Node)
//
// Capa fina: recibe los datos del recibo desde el frontend, los reenvía
// al microservicio Python (LibreOffice) que arma el PDF, y sube el
// resultado a Supabase Storage. Vercel nunca corre LibreOffice acá.

import { createClient } from "@supabase/supabase-js";

const GENERADOR_URL = process.env.RECIBOS_SERVICE_URL; // ej: https://recibos.fly.dev
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY // service_role: esta función corre server-side
);

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const datosRecibo = req.body; // { nombre_apellido, periodo_abonado, ... }

  try {
    const resp = await fetch(`${GENERADOR_URL}/generar-recibo`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(datosRecibo),
    });

    if (!resp.ok) {
      const detalle = await resp.text();
      return res.status(502).json({ error: "Fallo el generador de PDF", detalle });
    }

    const pdfBuffer = Buffer.from(await resp.arrayBuffer());
    const nombreArchivo = `recibos/${datosRecibo.cuil}_${datosRecibo.periodo_abonado}.pdf`;

    const { error: uploadError } = await supabase.storage
      .from("recibos")
      .upload(nombreArchivo, pdfBuffer, {
        contentType: "application/pdf",
        upsert: true,
      });

    if (uploadError) {
      return res.status(500).json({ error: "Fallo al subir a Supabase", detalle: uploadError.message });
    }

    const { data: urlData } = supabase.storage.from("recibos").getPublicUrl(nombreArchivo);

    return res.status(200).json({ url: urlData.publicUrl });
  } catch (err) {
    return res.status(500).json({ error: "Error inesperado", detalle: err.message });
  }
}
