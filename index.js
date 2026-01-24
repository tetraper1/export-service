app.get("/test-vw-guias", async (req, res) => {
  const pageSize = Number(req.query.pageSize || 100);
  const lastId = req.query.lastId || null;

  let sql = `
    select id, numero_guia, cuenta_cliente, fecha_de_creacion
    from public.vw_guias_report
    where 1 = 1
  `;
  const params = [];

  if (lastId) {
    params.push(lastId);
    sql += ` and id > $${params.length}`;
  }

  sql += `
    order by id asc
    limit ${pageSize}
  `;

  const { rows } = await pool.query(sql, params);

  res.json({
    count: rows.length,
    lastId: rows.length ? rows[rows.length - 1].id : null,
    rows
  });
});
