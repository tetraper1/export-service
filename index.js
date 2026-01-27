import express from "express";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import crypto from "crypto";

const app = express();
app.use(express.json());

// -------------------- Config (Starter-friendly) --------------------
const MAX_ROWS = 35_000;                  // tu requisito
const PAGE_SIZE = 5_000;                  // paginado por id (keyset)
const MAX_CONCURRENT_EXPORTS = Number(process.env.MAX_CONCURRENT_EXPORTS || 1); // Starter: 1 recomendado
const JOB_TTL_MS = Number(process.env.JOB_TTL_MS || 20 * 60 * 1000);            // 20 min
const MAX_PENDING_JOBS = Number(process.env.MAX_PENDING_JOBS || 200);

// -------------------- Supabase --------------------
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// -------------------- Columnas --------------------
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

const COLUMN_MAP = Object.fromEntries(COLUMNS.map(c => [c.key, c]));

const DATE_COLS = new Set([
  "fecha_de_creacion",
  "fecha_embarcado",
  "fecha_entregado",
  "fecha_recolectado",
  "fecha_ultima_actualizacion",
]);

// -------------------- Query builder (keyset pagination) --------------------
const buildQuery = (filters, selectColumns) => {
  let query = supabase
    .from("vw_guias_report")
    .select(selectColumns.join(","));

  if (filters.fecha_desde) query = query.gte("fecha_de_creacion", filters.fecha_desde);
  if (filters.fecha_hasta) query = query.lte("fecha_de_creacion", filters.fecha_hasta);

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

// -------------------- Simple queue (no deps) --------------------
let activeExports = 0;
const queue = [];

function runNext() {
  if (activeExports >= MAX_CONCURRENT_EXPORTS) return;
  const next = queue.shift();
  if (!next) return;

  activeExports++;
  next()
    .catch(() => {})
    .finally(() => {
      activeExports--;
      runNext();
    });
}

function enqueue(fn) {
  return new Promise((resolve, reject) => {
    queue.push(async () => {
      try {
        resolve(await fn());
      } catch (e) {
        reject(e);
      }
    });
    runNext();
  });
}

// -------------------- Jobs registry (in-memory) --------------------
const jobs = new Map();
const now = () => Date.now();

function makeId() {
  return crypto.randomUUID ? crypto.randomUUID() : crypto.randomBytes(16).toString("hex");
}

function cleanupJobs() {
  const t = now();
  for (const [jobId, job] of jobs) {
    if (t - job.updatedAt > JOB_TTL_MS) {
      if (job.filePath && fs.existsSync(job.filePath)) fs.unlink(job.filePath, () => {});
      jobs.delete(jobId);
    }
  }
}
setInterval(cleanupJobs, 60_000).unref();

// -------------------- Excel generation --------------------
async function generateExcelJob(jobId, filters) {
  const job = jobs.get(jobId);
  if (!job) return;

  job.status = "processing";
  job.updatedAt = now();

  const requestedColumns = Array.isArray(filters.columnas) ? filters.columnas : [];
  const finalColumnKeys =
    requestedColumns.length > 0
      ? requestedColumns.filter(col => COLUMN_MAP[col]).sort((a, b) => a.localeCompare(b))
      : COLUMNS.map(c => c.key);

  if (finalColumnKeys.length === 0) {
    job.status = "error";
    job.error = "No hay columnas válidas";
    job.updatedAt = now();
    return;
  }

  const excelColumns = finalColumnKeys.map((key) => {
    const base = COLUMN_MAP[key];
    if (DATE_COLS.has(key)) return { ...base, key, style: { numFmt: "yyyy-mm-dd hh:mm" } };
    return { ...base, key };
  });

  const selectColumns = ["id", ...finalColumnKeys];

  const fileName = `guias_export_${jobId}.xlsx`;
  const filePath = path.join("/tmp", fileName);

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: filePath,
    useStyles: true,
    useSharedStrings: false,
  });

  const sheet = workbook.addWorksheet("Guías");
  sheet.columns = excelColumns;

  // Define numFmt por columna (evita costo por celda)
  for (const key of DATE_COLS) {
    if (finalColumnKeys.includes(key)) sheet.getColumn(key).numFmt = "yyyy-mm-dd hh:mm";
  }

  let total = 0;
  let lastId = null;

  while (total < MAX_ROWS) {
    const remaining = MAX_ROWS - total;

    let query = buildQuery(filters, selectColumns)
      .order("id", { ascending: true })
      .limit(Math.min(PAGE_SIZE, remaining));

    if (lastId) query = query.gt("id", lastId);

    const { data, error } = await query;
    if (error) throw error;
    if (!data || data.length === 0) break;

    for (const row of data) {
      // Mutación ligera de fechas
      for (const k of DATE_COLS) if (row[k]) row[k] = new Date(row[k]);
      sheet.addRow(row).commit();
    }

    lastId = data[data.length - 1].id;
    total += data.length;

    job.rows = total;
    job.updatedAt = now();
  }

  await workbook.commit();

  job.status = "ready";
  job.rows = total;
  job.filePath = filePath;
  job.fileName = fileName;
  job.updatedAt = now();
}

// -------------------- Endpoints --------------------

// 1) Crear export (responde rápido, sin timeouts)
app.post("/export", async (req, res) => {
  try {
    if (jobs.size >= MAX_PENDING_JOBS) {
      return res.status(429).json({ error: "Demasiados exports en cola. Intenta más tarde." });
    }

    const filters = req.body || {};
    const jobId = makeId();

    jobs.set(jobId, {
      id: jobId,
      status: "queued", // queued | processing | ready | error
      createdAt: now(),
      updatedAt: now(),
      rows: 0,
      error: null,
      filePath: null,
      fileName: null,
    });

    enqueue(async () => {
      try {
        await generateExcelJob(jobId, filters);
      } catch (err) {
        const job = jobs.get(jobId);
        if (job) {
          job.status = "error";
          job.error = err?.message || "Error generando excel";
          job.updatedAt = now();
        }
      }
    }).catch(() => {});

    return res.status(202).json({
      ok: true,
      jobId,
      status_url: `/export/status/${jobId}`,
      download_url: `/export/download/${jobId}`,
      note: "Generación en background. Consulta status_url hasta que esté ready."
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
});

// 2) Consultar estado/progreso
app.get("/export/status/:jobId", (req, res) => {
  const { jobId } = req.params;
  const job = jobs.get(jobId);
  if (!job) return res.status(404).json({ error: "Job no encontrado" });

  return res.json({
    id: job.id,
    status: job.status,
    rows: job.rows,
    error: job.error,
    ready: job.status === "ready",
    download_url: job.status === "ready" ? `/export/download/${jobId}` : null,
    updatedAt: job.updatedAt,
  });
});

// 3) Descargar (streaming del archivo local)
app.get("/export/download/:jobId", (req, res) => {
  const { jobId } = req.params;
  const job = jobs.get(jobId);
  if (!job) return res.status(404).json({ error: "Job no encontrado" });

  if (job.status !== "ready" || !job.filePath) {
    return res.status(409).json({ error: "Archivo aún no está listo", status: job.status });
  }

  if (!fs.existsSync(job.filePath)) {
    return res.status(404).json({ error: "Archivo no encontrado en disco (expiró o se limpió)" });
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${job.fileName}"`);
  res.setHeader("X-Accel-Buffering", "no");

  const stream = fs.createReadStream(job.filePath);

  stream.on("error", (e) => {
    console.error("read stream error", e);
    if (!res.headersSent) res.status(500).json({ error: "Error leyendo archivo" });
    else res.end();
  });

  // No borramos al final para permitir reintentos; se limpia por TTL.
  stream.pipe(res);
});

// -------------------- Server tuning --------------------
const PORT = process.env.PORT || 3000;
const server = app.listen(PORT, () => {
  console.log(`Export service running on port ${PORT}`);
});

// Evita timeouts del server Node (la plataforma puede imponer límites)
server.requestTimeout = 0;
server.headersTimeout = 0;
server.keepAliveTimeout = 65_000;
