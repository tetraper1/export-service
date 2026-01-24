import express from "express";
import { createClient } from "@supabase/supabase-js";

const app = express();
app.use(express.json());

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

app.get("/export/test", async (req, res) => {
  try {
    const PAGE_SIZE = 2000;
    let lastId = null;
    let total = 0;

    while (true) {
      let query = supabase
        .from("vw_guias_report")
        .select("id")
        .order("id", { ascending: true })
        .limit(PAGE_SIZE);

      if (lastId) {
        query = query.gt("id", lastId);
      }

      const { data, error } = await query;

      if (error) throw error;
      if (!data || data.length === 0) break;

      lastId = data[data.length - 1].id;
      total += data.length;

      console.log(`Página procesada. Total filas: ${total}`);
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

app.listen(3000, () => {
  console.log("Export test running on port 3000");
});
