import express from "express";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

const app = express();
app.use(express.json());

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

    // 📘 Workbook en streaming
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: filePath
    });

    const sheet = workbook.addWorksheet("Guías");

    // 📌 Definir columnas
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

      // filtros
      if (filters.fecha_desde) {
        query = query.gte(
          "fecha_de_creacion",
          filters.fecha_desde
        );
      }

      if (filters.fecha_hasta) {
        query = query.lte(
          "fecha_de_creacion",
          filters.fecha_hasta
        );
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
