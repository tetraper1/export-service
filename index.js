import express from "express";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const app = express();
app.use(express.json());

// Supabase client
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// Columnas permitidas para export
const COLUMNS = [
  { header: "calle_destino", key: "calle_destino" },
  { header: "calle_remitente_origen", key: "calle_remitente_origen" },
  { header: "cintillo_de_seguridad", key: "cintillo_de_seguridad" },
  { header: "codigo_postal_destino", key: "codigo_postal_destino" },
  { header: "codigo_postal_remitente_origen", key: "codigo_postal_remitente_origen" },
  { header: "codigo_rastreo", key: "codigo_rastreo" },
  { header: "colonia_destino", key: "colonia_destino" },
  { header: "colonia_remitente_origen", key: "colonia_remitente_origen" },
  { header: "comentarios_referencias", key: "comentarios_referencias" },
  { header: "contenido_del_paquete", key: "contenido_del_paquete" },
  { header: "creador", key: "creador" },
  { header: "cuenta_cliente", key: "cuenta_cliente" },
  { header: "estado_destino", key: "estado_destino" },
  { header: "estado_remitente_origen", key: "estado_remitente_origen" },
  { header: "evento_actual", key: "evento_actual" },
  { header: "evento_actual_id_nivel", key: "evento_actual_id_nivel" },
  { header: "fecha_de_creacion", key: "fecha_de_creacion" },
  { header: "fecha_embarcado", key: "fecha_embarcado" },
  { header: "fecha_entregado", key: "fecha_entregado" },
  { header: "fecha_recolectado", key: "fecha_recolectado" },
  { header: "fecha_ultima_actualizacion", key: "fecha_ultima_actualizacion" },
  { header: "guia_madre", key: "guia_madre" },
  { header: "lote_de_guias", key: "lote_de_guias" },
  { header: "manifiesto_de_enlace", key: "manifiesto_de_enlace" },
  { header: "manifiesto_de_recoleccion", key: "manifiesto_de_recoleccion" },
  { header: "manifiesto_de_ruta", key: "manifiesto_de_ruta" },
  { header: "medida_alto", key: "medida_alto" },
  { header: "medida_ancho", key: "medida_ancho" },
  { header: "medida_largo", key: "medida_largo" },
  { header: "municipio_destino", key: "municipio_destino" },
  { header: "municipio_remitente_origen", key: "municipio_remitente_origen" },
  { header: "nombre_de_quien_recibe_receptor", key: "nombre_de_quien_recibe_receptor" },
  { header: "nombre_razon_social_destino", key: "nombre_razon_social_destino" },
  { header: "nombre_razon_social_remitente_origen", key: "nombre_razon_social_remitente_origen" },
  { header: "nombre_ruta_entrega", key: "nombre_ruta_entrega" },
  { header: "nombre_ruta_recoleccion", key: "nombre_ruta_recoleccion" },
  { header: "numero_de_guia", key: "numero_de_guia" },
  { header: "peso_unitario", key: "peso_unitario" },
  { header: "peso_volumetrico_unitario", key: "peso_volumetrico_unitario" },
  { header: "telefono_1_destino", key: "telefono_1_destino" },
  { header: "telefono_2_destino", key: "telefono_2_destino" },
  { header: "tipo_de_embalaje", key: "tipo_de_embalaje" },
  { header: "ultima_sucursal", key: "ultima_sucursal" },
  { header: "ultimo_operador_asignado", key: "ultimo_operador_asignado" },
  { header: "usuario_confirma_entrega", key: "usuario_confirma_entrega" },
  { header: "valor_declarado", key: "valor_declarado" }
];

// Índice rápido
const COLUMN_MAP = Object.fromEntries(
  COLUMNS.map(col => [col.key, col])
);

// Builder de filtros (NO paginación aquí)
const buildQuery = (filters, selectColumns) => {
  let query = supabase
    .from("vw_guias_report")
    .select(selectColumns.join(","));

  if (filters.fecha_desde) {
    query = query.gte("fecha_de_creacion", filters.fecha_desde);
  }

  if (filters.fecha_hasta) {
    query = query.lte("fecha_de_creacion", filters.fecha_hasta);
  }

  if (Array.isArray(filters.numeros_guia) && filters.numeros_guia.length) {
    query = query.in("numero_de_guia", filters.numeros_guia);
  }

  if (Array.isArray(filters.codigos_rastreo) && filters.codigos_rastreo.length) {
    query = query.in("codigo_rastreo", filters.codigos_rastreo);
  }

  if (Array.isArray(filters.guias_madre) && filters.guias_madre.length) {
    query = query.in("guia_madre", filters.guias_madre);
  }

  if (Array.isArray(filters.cuenta_cliente_id) && filters.cuenta_cliente_id.length) {
    query = query.in("cuenta_cliente_id", filters.cuenta_cliente_id);
  }

  if (Array.isArray(filters.ultima_sucursal) && filters.ultima_sucursal.length) {
    query = query.in("ultima_sucursal", filters.ultima_sucursal);
  }

  if (Array.isArray(filters.nombres_ruta_entrega) && filters.nombres_ruta_entrega.length) {
    query = query.in("nombre_ruta_entrega", filters.nombres_ruta_entrega);
  }

  if (Array.isArray(filters.evento_actual_id_nivel) && filters.evento_actual_id_nivel.length) {
    query = query.in("evento_actual_id_nivel", filters.evento_actual_id_nivel);
  }

  return query;
};

const DATE_COLUMNS = new Set([
  "fecha_de_creacion",
  "fecha_ultima_actualizacion",
  "fecha_embarcado",
  "fecha_entregado",
  "fecha_recolectado"
]);


// Endpoint export
app.post("/export/test", async (req, res) => {
  try {
    const filters = req.body;
    const PAGE_SIZE = 2000;
    let lastId = null;
    let total = 0;

    // Columnas solicitadas
    const requestedColumns = Array.isArray(filters.columnas)
      ? filters.columnas
      : [];

    const finalColumnKeys =
      requestedColumns.length > 0
        ? requestedColumns
            .filter(col => COLUMN_MAP[col])
            .sort((a, b) => a.localeCompare(b))
        : COLUMNS.map(c => c.key);

    if (finalColumnKeys.length === 0) {
      return res.status(400).json({ error: "No hay columnas válidas" });
    }

    const excelColumns = finalColumnKeys.map(key => {
  const col = { ...COLUMN_MAP[key] };

  if (DATE_COLUMNS.has(key)) {
    col.style = {
      numFmt: "yyyy-mm-dd hh:mm"
    };
  }

  return col;
});
    sheet.columns = excelColumns;
    const selectColumns = ["id", ...finalColumnKeys];

    const filePath = path.join("/tmp", `guias_export_${Date.now()}.xlsx`);

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: filePath
    });

    const sheet = workbook.addWorksheet("Guías");
    sheet.columns = excelColumns;

    while (true) {
      let query = buildQuery(filters, selectColumns)
        .order("id", { ascending: true })
        .limit(PAGE_SIZE);

      if (lastId) query = query.gt("id", lastId);

      const { data, error } = await query;
      if (error) throw error;
      if (!data || data.length === 0) break;

      for (const row of data) {
        const excelRow = { ...row };

for (const key of DATE_COLUMNS) {
  if (excelRow[key]) {
    excelRow[key] = new Date(excelRow[key]);
  }
}

sheet.addRow(excelRow).commit();
      }

      lastId = data[data.length - 1].id;
      total += data.length;
    }

    await workbook.commit();

    return res.json({
      ok: true,
      filas_exportadas: total,
      columnas: finalColumnKeys,
      download_url: `/export/download?file=${path.basename(filePath)}`
    });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
});

// Descarga
app.get("/export/download", (req, res) => {
  const { file } = req.query;
  if (!file) return res.status(400).json({ error: "file requerido" });

  const filePath = path.join("/tmp", file);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: "Archivo no encontrado" });
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    `attachment; filename="${file}"`
  );

  const stream = fs.createReadStream(filePath);
  stream.pipe(res);
  stream.on("end", () => fs.unlink(filePath, () => {}));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Export service running on port ${PORT}`);
});
