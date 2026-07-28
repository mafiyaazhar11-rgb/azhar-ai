// ══ PRICE DISCREPANCY TRACKER ════════════════════════════════════════════
// Every team member uploads their own weekly log of System Price vs LPO Price
// mismatches they've found (self-reported, same pattern as CS Workload) —
// logged under their own login, so ownership is never ambiguous and nobody
// can upload claiming to be someone else.
//
// The whole point of this module is catching REPEATS: if the same
// Customer + SKU discrepancy shows up again in a later week while still
// unresolved, that's a signal the team leader isn't fixing it — so instead of
// storing a "repeat" flag at write time, every read groups rows by a
// normalized Customer+SKU key and works out at query time: how many times has
// this been reported, is it still open, and how many days has it been open
// since it was first flagged. That's what powers the executive escalation
// view — Repeated + still unresolved + days-open, sorted worst-first.
const XLSX = require('xlsx');
const crypto = require('crypto');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog) {

  (async function initTables() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS price_discrepancies (
        id SERIAL PRIMARY KEY,
        reported_by_username TEXT NOT NULL,
        reported_by_full_name TEXT,
        week_ending DATE,
        customer_name TEXT NOT NULL,
        sku_code TEXT NOT NULL,
        sku_description TEXT,
        system_price NUMERIC,
        lpo_price NUMERIC,
        discrepancy NUMERIC,
        remarks TEXT,
        status TEXT NOT NULL DEFAULT 'Open',  -- 'Open' or 'Resolved'
        resolved_at TIMESTAMPTZ,
        resolved_by TEXT,
        file_name TEXT,
        batch_id TEXT,
        uploaded_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pdisc_cust ON price_discrepancies (customer_name)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pdisc_sku ON price_discrepancies (sku_code)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pdisc_rep ON price_discrepancies (reported_by_username)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pdisc_batch ON price_discrepancies (batch_id)`);
      console.log('Price Discrepancy module DB ready');
    } catch (e) { console.error('Price Discrepancy init error:', e.message); }
  })();

  function normalizeHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
  function fc(keys, names) {
    return keys.find(function (k) {
      var nk = normalizeHeader(k);
      return names.some(function (n) { return nk.indexOf(n) !== -1; });
    }) || null;
  }
  // Normalized grouping key — collapses whitespace/case/punctuation differences
  // so "Carrefour Mall Of Emirates" and "carrefour  mall of emirates" (or a
  // SKU with/without dashes) are recognized as the same discrepancy thread.
  function custKey(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function skuKey(s) { return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''); }

  // ── Upload: each team member uploads their own weekly file. Identity comes
  // from the logged-in user, not from anything in the file. Always inserted as
  // a NEW row (this is a weekly log, not an update-in-place like CS Workload) —
  // repeats across weeks are exactly the signal this module is built to catch. ──
  app.post('/api/price-discrepancy/upload', requireAuth, upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
      catch (e) { return res.status(400).json({ error: 'Could not read that file. Upload an .xlsx/.csv.' }); }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      var raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!raw.length) return res.status(400).json({ error: 'No rows found in that file.' });

      var keys = Object.keys(raw[0]);
      var C = {
        week: fc(keys, ['weekending']) || fc(keys, ['week']),
        customer: fc(keys, ['customername']) || fc(keys, ['customer']),
        sku: fc(keys, ['skucode']) || fc(keys, ['itemcode']) || fc(keys, ['sku']),
        desc: fc(keys, ['skudescription']) || fc(keys, ['description']),
        system: fc(keys, ['systemprice']) || fc(keys, ['system']),
        lpo: fc(keys, ['lpoprice']) || fc(keys, ['lpo']),
        remarks: fc(keys, ['remarks', 'reason'])
      };
      if (!C.customer) return res.status(400).json({ error: 'Could not find a Customer Name column in that file.' });
      if (!C.sku) return res.status(400).json({ error: 'Could not find a SKU Code column in that file.' });
      if (!C.system || !C.lpo) return res.status(400).json({ error: 'Could not find both System Price and LPO Price columns in that file.' });

      var repUsername = req.user.username;
      var repFullName = req.user.full_name || req.user.username;
      var batchId = crypto.randomUUID();
      var inserted = 0, skipped = 0;

      for (var i = 0; i < raw.length; i++) {
        var row = raw[i];
        var customer = String(row[C.customer] || '').trim();
        var sku = String(row[C.sku] || '').trim();
        if (!customer || !sku) { skipped++; continue; }
        var systemPrice = row[C.system] === '' ? null : +row[C.system];
        var lpoPrice = row[C.lpo] === '' ? null : +row[C.lpo];
        if (systemPrice == null || lpoPrice == null || isNaN(systemPrice) || isNaN(lpoPrice)) { skipped++; continue; }
        // Only log genuine mismatches — an identical price isn't a discrepancy.
        if (Math.round((systemPrice - lpoPrice) * 100) === 0) { skipped++; continue; }
        var weekEnding = C.week ? row[C.week] : null;
        var desc = C.desc ? String(row[C.desc] || '').trim() : null;
        var remarks = C.remarks ? String(row[C.remarks] || '').trim() : null;

        await pool.query(
          `INSERT INTO price_discrepancies
             (reported_by_username, reported_by_full_name, week_ending, customer_name, sku_code,
              sku_description, system_price, lpo_price, discrepancy, remarks, file_name, batch_id)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
          [repUsername, repFullName, weekEnding, customer, sku, desc, systemPrice, lpoPrice,
           Math.round((lpoPrice - systemPrice) * 100) / 100, remarks, req.file.originalname, batchId]
        );
        inserted++;
      }

      auditLog(req.user.uid, repUsername, 'PRICE_DISCREPANCY_UPLOAD',
        'Uploaded ' + inserted + ' discrepancy row(s), skipped ' + skipped + ' (' + req.file.originalname + ')',
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, inserted: inserted, skipped: skipped, batchId: batchId });
    } catch (e) {
      console.error('price-discrepancy upload error:', e.message);
      res.status(500).json({ error: e.message });
    }
  });

  // ── Manual single-row add (for someone without a file, or a one-off catch) ──
  app.post('/api/price-discrepancy/add', requireAuth, async function (req, res) {
    try {
      var b = req.body || {};
      var customer = String(b.customerName || '').trim();
      var sku = String(b.skuCode || '').trim();
      var systemPrice = +b.systemPrice, lpoPrice = +b.lpoPrice;
      if (!customer || !sku) return res.status(400).json({ error: 'Customer Name and SKU Code are required.' });
      if (isNaN(systemPrice) || isNaN(lpoPrice)) return res.status(400).json({ error: 'System Price and LPO Price must be numbers.' });
      var r = await pool.query(
        `INSERT INTO price_discrepancies
           (reported_by_username, reported_by_full_name, week_ending, customer_name, sku_code,
            sku_description, system_price, lpo_price, discrepancy, remarks)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) RETURNING id`,
        [req.user.username, req.user.full_name || req.user.username, b.weekEnding || null, customer, sku,
         b.skuDescription || null, systemPrice, lpoPrice, Math.round((lpoPrice - systemPrice) * 100) / 100, b.remarks || null]
      );
      auditLog(req.user.uid, req.user.username, 'PRICE_DISCREPANCY_ADD', 'Manually logged ' + customer + ' / ' + sku, req.headers['x-forwarded-for'] || '');
      res.json({ success: true, id: r.rows[0].id });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Manage Uploads: list recent upload batches for deletion of a single
  // wrong file, without a blanket Clear All. Superadmin/subadmin only. ──
  app.get('/api/price-discrepancy/batches', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query(
        `SELECT batch_id, reported_by_username, reported_by_full_name, file_name,
                MIN(uploaded_at) AS uploaded_at, COUNT(*) AS row_count
         FROM price_discrepancies
         WHERE batch_id IS NOT NULL
         GROUP BY batch_id, reported_by_username, reported_by_full_name, file_name
         ORDER BY MIN(uploaded_at) DESC
         LIMIT 100`
      );
      res.json({ success: true, batches: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.delete('/api/price-discrepancy/batches/:batchId', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var batchId = req.params.batchId;
      var check = await pool.query(
        `SELECT reported_by_full_name, file_name, COUNT(*) AS cnt FROM price_discrepancies WHERE batch_id = $1 GROUP BY reported_by_full_name, file_name`,
        [batchId]
      );
      if (!check.rows.length) return res.status(404).json({ error: 'That upload was not found — it may have already been deleted.' });
      var r = await pool.query('DELETE FROM price_discrepancies WHERE batch_id = $1', [batchId]);
      var info = check.rows[0];
      await auditLog(req.user.uid, req.user.username, 'PRICE_DISCREPANCY_BATCH_DELETE',
        'Deleted upload batch (' + r.rowCount + ' rows) \u2014 ' + info.reported_by_full_name + ' / ' + info.file_name,
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, deleted: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Core grouped query: every row, joined to its Customer+SKU thread stats
  // (times reported, still open, first-reported date, days unresolved). This
  // single query powers the consolidated view, the escalation view, and both
  // export endpoints, so the "repeated" definition never drifts between them. ──
  async function fetchGrouped(whereSql, params) {
    var sql = `
      WITH keyed AS (
        SELECT *,
          lower(regexp_replace(trim(customer_name), '\\s+', ' ', 'g')) AS cust_key,
          lower(regexp_replace(trim(sku_code), '[^a-zA-Z0-9]', '', 'g')) AS sku_key
        FROM price_discrepancies
      ),
      agg AS (
        SELECT cust_key, sku_key,
          COUNT(*) AS times_reported,
          COUNT(DISTINCT week_ending) AS distinct_weeks,
          MIN(uploaded_at) AS first_reported_at,
          MAX(uploaded_at) AS last_reported_at,
          BOOL_OR(status = 'Open') AS still_open
        FROM keyed GROUP BY cust_key, sku_key
      )
      SELECT k.*, a.times_reported, a.distinct_weeks, a.first_reported_at, a.last_reported_at, a.still_open,
        (a.times_reported > 1 AND a.still_open) AS is_repeated,
        CASE WHEN a.still_open THEN EXTRACT(DAY FROM (NOW() - a.first_reported_at))::int ELSE NULL END AS days_unresolved
      FROM keyed k JOIN agg a ON k.cust_key = a.cust_key AND k.sku_key = a.sku_key
      ${whereSql || ''}
      ORDER BY days_unresolved DESC NULLS LAST, k.uploaded_at DESC`;
    return pool.query(sql, params || []);
  }

  // ── Consolidated view — superadmin/subadmin, all reps, filterable by rep/status ──
  app.get('/api/price-discrepancy/consolidated', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.rep_username) { params.push(req.query.rep_username); clauses.push('k.reported_by_username = $' + params.length); }
      if (req.query.status) { params.push(req.query.status); clauses.push('k.status = $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var r = await fetchGrouped(where, params);
      var reps = await pool.query(`SELECT DISTINCT reported_by_username, reported_by_full_name FROM price_discrepancies ORDER BY reported_by_full_name`);
      res.json({ success: true, rows: r.rows, reps: reps.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Escalation view — repeated AND still unresolved, worst-first. This is
  // the exact list that should go in front of the executive. ──
  app.get('/api/price-discrepancy/escalations', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var threshold = req.query.days_threshold ? +req.query.days_threshold : 7;
      // is_repeated/days_unresolved are computed columns, so filtering happens
      // in JS after fetching rather than in SQL WHERE.
      var all = await fetchGrouped('', []);
      var seen = {}, list = [];
      all.rows.forEach(function (row) {
        if (!row.is_repeated || !row.still_open) return;
        var key = row.cust_key + '|' + row.sku_key;
        if (seen[key]) return; // one line per thread, not per occurrence
        seen[key] = true;
        list.push(row);
      });
      list = list.filter(function (row) { return row.days_unresolved >= threshold; });
      list.sort(function (a, b) { return b.days_unresolved - a.days_unresolved; });
      res.json({ success: true, rows: list, threshold: threshold });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Resolve: closes the WHOLE thread (every Open row sharing the same
  // normalized Customer+SKU key), not just the one row clicked — so the days-
  // unresolved clock actually stops instead of restarting on the next report. ──
  app.patch('/api/price-discrepancy/:id/resolve', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var row = await pool.query('SELECT customer_name, sku_code FROM price_discrepancies WHERE id = $1', [req.params.id]);
      if (!row.rows.length) return res.status(404).json({ error: 'Not found.' });
      var ck = custKey(row.rows[0].customer_name), sk = skuKey(row.rows[0].sku_code);
      var r = await pool.query(
        `UPDATE price_discrepancies
         SET status = 'Resolved', resolved_at = NOW(), resolved_by = $1
         WHERE status = 'Open'
           AND lower(regexp_replace(trim(customer_name), '\\s+', ' ', 'g')) = $2
           AND lower(regexp_replace(trim(sku_code), '[^a-zA-Z0-9]', '', 'g')) = $3`,
        [req.user.username, ck, sk]
      );
      await auditLog(req.user.uid, req.user.username, 'PRICE_DISCREPANCY_RESOLVE',
        'Resolved ' + r.rowCount + ' row(s) \u2014 ' + row.rows[0].customer_name + ' / ' + row.rows[0].sku_code,
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, resolved: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── A user's own submissions only ──
  app.get('/api/price-discrepancy/my-entries', requireAuth, async function (req, res) {
    try {
      var r = await fetchGrouped('WHERE k.reported_by_username = $1', [req.user.username]);
      res.json({ success: true, rows: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Shared export builder — one sheet of detail rows + one Executive
  // Escalation sheet (repeated + still-open threads only, one line each). ──
  async function buildWorkbook(detailRows, scopeLabel) {
    var ExcelJS = require('exceljs');
    var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', GOLD = 'FFFCF3CF', RED = 'FFFADBD8';
    function headerRow(ws, cells) {
      var row = ws.addRow(cells);
      row.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      return row;
    }

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    var ds = wb.addWorksheet('Discrepancy Detail');
    ds.columns = [{ width: 14 }, { width: 22 }, { width: 30 }, { width: 14 }, { width: 34 },
                  { width: 12 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 12 }, { width: 28 }];
    var t1 = ds.addRow(['PRICE DISCREPANCY LOG \u2014 ' + scopeLabel]);
    ds.mergeCells('A' + t1.number + ':K' + t1.number);
    t1.font = { bold: true, size: 13, color: { argb: NAVY } };
    ds.addRow(['Generated', new Date().toLocaleString('en-AE')]);
    ds.addRow([]);
    headerRow(ds, ['Week Ending', 'Reported By', 'Customer Name', 'SKU Code', 'Description',
                   'System Price', 'LPO Price', 'Discrepancy', 'Status', 'Times Reported', 'Days Unresolved']);
    detailRows.forEach(function (row) {
      var r = ds.addRow([
        row.week_ending ? new Date(row.week_ending).toLocaleDateString('en-AE') : '',
        row.reported_by_full_name || row.reported_by_username,
        row.customer_name, row.sku_code, row.sku_description || '',
        +row.system_price, +row.lpo_price, +row.discrepancy, row.status,
        +row.times_reported, row.days_unresolved == null ? '' : +row.days_unresolved
      ]);
      if (row.is_repeated && row.still_open) r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } }; });
    });

    var seen = {}, escalations = [];
    detailRows.forEach(function (row) {
      if (!row.is_repeated || !row.still_open) return;
      var key = row.cust_key + '|' + row.sku_key;
      if (seen[key]) return;
      seen[key] = true;
      escalations.push(row);
    });
    escalations.sort(function (a, b) { return b.days_unresolved - a.days_unresolved; });

    var es = wb.addWorksheet('Executive Escalation');
    es.columns = [{ width: 30 }, { width: 14 }, { width: 22 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }];
    var t2 = es.addRow(['UNRESOLVED & REPEATED \u2014 FOR EXECUTIVE REVIEW']);
    es.mergeCells('A' + t2.number + ':G' + t2.number);
    t2.font = { bold: true, size: 13, color: { argb: NAVY } };
    es.addRow(['Generated', new Date().toLocaleString('en-AE')]);
    es.addRow([]);
    headerRow(es, ['Customer Name', 'SKU Code', 'Last Reported By', 'System Price', 'LPO Price', 'Times Reported', 'Days Unresolved']);
    if (!escalations.length) {
      es.addRow(['No repeated, unresolved discrepancies right now.']);
    } else {
      escalations.forEach(function (row) {
        var r = es.addRow([row.customer_name, row.sku_code, row.reported_by_full_name || row.reported_by_username,
          +row.system_price, +row.lpo_price, +row.times_reported, +row.days_unresolved]);
        r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
      });
    }
    return wb;
  }

  // ── Superadmin/subadmin: consolidated download, optionally scoped to one rep ──
  app.get('/api/price-discrepancy/consolidated/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.rep_username) { params.push(req.query.rep_username); clauses.push('k.reported_by_username = $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var r = await fetchGrouped(where, params);
      var label = req.query.rep_username ? ('Team Member: ' + req.query.rep_username) : 'All Team Members';
      var wb = await buildWorkbook(r.rows, label);
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Price_Discrepancy_Consolidated_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('price-discrepancy consolidated export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });

  // ── Any logged-in user: their own submissions, downloadable by themselves ──
  app.get('/api/price-discrepancy/my-entries/export', requireAuth, async function (req, res) {
    try {
      var r = await fetchGrouped('WHERE k.reported_by_username = $1', [req.user.username]);
      var wb = await buildWorkbook(r.rows, (req.user.full_name || req.user.username));
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="My_Price_Discrepancies_' + req.user.username + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('price-discrepancy my-export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });
};
