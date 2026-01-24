import express from "express";
import { createClient } from "@supabase/supabase-js";

const app = express();
app.use(express.json());

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

// ===============================
// ENDPOINT DE PRUEBA CON FILTROS
// ===============================
app.post("/export/test", async (req, res) => {
  try {
    const {
      fecha_desde,
      fecha_hasta,
      cuenta_cliente_id,
      estado_destino,
      evento_actual_id_nivel
    } = req.body;

    const PAGE_SIZE = 2000;
    let lastId = null;
    let total = 0;

    while (true) {
      let query = supabase
        .from("vw_guias_report")
        .select("id")
        .order("id", { ascending: true })
        .limit(PAGE_SIZE);

      // 🔑 keyset pagination
      if (lastId) {
        query = query.gt("id", lastId);
      }

      // 🧩 filtros
      if (fecha_desde) {
        query = query.gte("fecha_de_creacion", fecha_desde);
      }

      if (fecha_hasta) {
        query = query.lte("fecha_de_creacion", fecha_hasta);
      }

      if (cuenta_cliente_id) {
        query = query.eq("cuenta_cliente_id", cuenta_cliente_id);
      }

      if (estado_destino) {
        query = query.eq("estado_destino", estado_destino);
      }

      if (evento_actual_id_nivel) {
        query = query.eq(
          "evento_actual_id_nivel",
          evento_actual_id_nivel
        );
      }

      const { data, error } = await query;

      if (error) throw error;
      if (!data || data.length === 0) break;

      lastId = data[data.length - 1].id;
      total += data.length;

      console.log(
        `Página OK | Último ID: ${lastId} | Total: ${total}`
      );
    }

    res.json({
      ok: true,
      filas_procesadas: total
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ===============================
// START SERVER
// ===============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Export test running on port ${PORT}`);
});
