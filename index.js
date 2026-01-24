import express from "express";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";
import path from "path";

const app = express();
app.use(express.json());

// 🔐 Supabase client (OBLIGATORIO)
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// 📌 Columnas del XLSX (OBLIGATORIO)
const COLUMNS = [
  { header: "ID", key: "id" },
  { header: "Número Guía", key: "numero_guia" },
  { header: "Código Rastreo", key: "codigo_rastreo" },
  { header: "Guía Madre", key: "guia_madre" },
  { header: "Ruta Entrega", key: "nombre_ruta_entrega" },
  { header: "Calle Destino", key: "calle_destino" },
  { header: "Cliente", key: "cuenta_cliente" },
  { header: "Cliente ID", key: "cuenta_cliente_id" },
  { header: "Estado Destino", key: "estado_destino" },
  { header: "Evento Actual", key: "evento_actual" },
  { header: "Fecha Creación", key: "fecha_de_creacion" }
];

// 🚀 Endpoint de prueba
app.post("/export/test", async (req, res) => {
  try {
    const filters = req.body;

    const PAGE_SIZE = 2000;
    let lastId = null;
    let total = 0;

    const filePath = path.join(
      "/tmp",
      `guias_export_${Date.now()}.xlsx`
    );

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: filePath
    });

    const sheet = workbook.addWorksheet("Guías");
    sheet.columns = COLUMNS;

    while (true) {
      let query = supabase
        .from("vw_guias_report")
        .select(COLUMNS.map(c => c.key).join(","))
        .order("id", { ascending: true })
        .limit(PAGE_SIZE);

      if (lastId) {
        query = query.gt("id", lastId);
      }

      if (filters.fecha_desde) {
        query = query.gte("fecha_de_creacion", filters.fecha_desde);
      }

      if (filters.fecha_hasta) {
        query = query.lte("fecha_de_creacion", filters.fecha_hasta);
      }

      if (filters.cuenta_cliente_id) {
        query = query.eq(
          "cuenta_cliente_id",
          filters.cuenta_cliente_id
        );
      }

      const { data, error } = await query;
      if (error) throw error;
      if (!data || data.length === 0) break;

      for (const row of data) {
        sheet.addRow(row).commit();
      }

      lastId = data[data.length - 1].id;
      total += data.length;

      console.log(`XLSX rows: ${total}`);
    }

    await workbook.commit();

    res.json({
      ok: true,
      filas_exportadas: total,
      file: filePath
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// 🌐 Render REQUIERE esto
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Export service running on port ${PORT}`);
});
