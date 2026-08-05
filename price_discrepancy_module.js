// ══ PRICE DISCREPANCY TRACKER ════════════════════════════════════════════
// Every team member uploads their own weekly log of System Price vs LPO Price
// mismatches they've found (self-reported, same pattern as CS Workload) —
// logged under their own login, so ownership is never ambiguous and nobody
// can upload claiming to be someone else.
//
// Re-upload behaviour (this is the whole point of the module): if a Customer +
// SKU discrepancy is still OPEN, seeing it again next week does NOT create a
// second row — the existing row is kept as-is (its original first_reported_at
// never moves), only its "last confirmed" timestamp and times_reported counter
// update, so Days Unresolved keeps counting from when it was truly first
// flagged. Only a genuinely new Customer+SKU (or one that recurs after being
// marked Resolved) becomes a new entry. That's what "Repeated" and the
// executive escalation view are built on.
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
      // first_reported_at: set once when the row is first created, NEVER
      // touched again — this is what Days Unresolved counts from.
      // times_reported: how many separate weekly uploads have re-confirmed
      // this same still-open issue (starts at 1).
      await pool.query(`ALTER TABLE price_discrepancies ADD COLUMN IF NOT EXISTS first_reported_at TIMESTAMPTZ`);
      await pool.query(`ALTER TABLE price_discrepancies ADD COLUMN IF NOT EXISTS times_reported INT DEFAULT 1`);
      // Who flagged this Customer+SKU issue THE FIRST TIME it was ever reported — kept
      // separate from reported_by_* (which is really "who last confirmed it's still open"
      // and gets overwritten on every re-report). Without this, once a second person
      // re-uploads the same still-open issue, the original reporter's name is lost.
      await pool.query(`ALTER TABLE price_discrepancies ADD COLUMN IF NOT EXISTS first_reported_by_username TEXT`);
      await pool.query(`ALTER TABLE price_discrepancies ADD COLUMN IF NOT EXISTS first_reported_by_full_name TEXT`);
      await pool.query(`UPDATE price_discrepancies SET first_reported_by_username = reported_by_username WHERE first_reported_by_username IS NULL`);
      await pool.query(`UPDATE price_discrepancies SET first_reported_by_full_name = reported_by_full_name WHERE first_reported_by_full_name IS NULL`);
      // Backfill for rows that existed before these columns did.
      await pool.query(`UPDATE price_discrepancies SET first_reported_at = uploaded_at WHERE first_reported_at IS NULL`);
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
  // Normalized matching key — collapses whitespace/case/punctuation differences
  // so "Carrefour Mall Of Emirates" and "carrefour  mall of emirates" (or a
  // SKU with/without dashes) are recognized as the same ongoing issue.
  function custKey(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  function skuKey(s) { return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''); }

  // Finds an existing OPEN row for this Customer+SKU, if any — used by both
  // upload and manual-add so re-confirming an open issue never duplicates it.
  async function findOpenMatch(customer, sku) {
    var r = await pool.query(
      `SELECT id, times_reported FROM price_discrepancies
       WHERE status = 'Open'
         AND lower(regexp_replace(trim(customer_name), '\\s+', ' ', 'g')) = $1
         AND lower(regexp_replace(trim(sku_code), '[^a-zA-Z0-9]', '', 'g')) = $2
       LIMIT 1`,
      [custKey(customer), skuKey(sku)]
    );
    return r.rows[0] || null;
  }

  // ── Upload: each team member uploads their own weekly file. Identity comes
  // from the logged-in user, not from anything in the file. If a row's
  // Customer+SKU already has an OPEN entry, that entry is refreshed in place
  // (latest prices/remarks, times_reported+1) rather than duplicated — a
  // genuinely new Customer+SKU becomes a new row. ──
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
      var inserted = 0, kept = 0, skipped = 0;

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
        var discrepancy = Math.round((lpoPrice - systemPrice) * 100) / 100;

        var existing = await findOpenMatch(customer, sku);
        if (existing) {
          // Same still-open issue seen again — refresh it in place, don't duplicate.
          // first_reported_at is untouched on purpose, so Days Unresolved keeps counting.
          await pool.query(
            `UPDATE price_discrepancies
             SET week_ending=$1, sku_description=COALESCE($2, sku_description), system_price=$3, lpo_price=$4,
                 discrepancy=$5, remarks=COALESCE($6, remarks), file_name=$7, batch_id=$8, uploaded_at=NOW(),
                 reported_by_username=$9, reported_by_full_name=$10, times_reported = times_reported + 1
             WHERE id=$11`,
            [weekEnding, desc, systemPrice, lpoPrice, discrepancy, remarks, req.file.originalname, batchId,
             repUsername, repFullName, existing.id]
          );
          kept++;
          continue;
        }

        await pool.query(
          `INSERT INTO price_discrepancies
             (reported_by_username, reported_by_full_name, first_reported_by_username, first_reported_by_full_name,
              week_ending, customer_name, sku_code,
              sku_description, system_price, lpo_price, discrepancy, remarks, file_name, batch_id,
              first_reported_at, times_reported)
           VALUES ($1,$2,$1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,NOW(),1)`,
          [repUsername, repFullName, weekEnding, customer, sku, desc, systemPrice, lpoPrice,
           discrepancy, remarks, req.file.originalname, batchId]
        );
        inserted++;
      }

      auditLog(req.user.uid, repUsername, 'PRICE_DISCREPANCY_UPLOAD',
        'Uploaded: ' + inserted + ' new, ' + kept + ' existing kept (days still counting), skipped ' + skipped + ' (' + req.file.originalname + ')',
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, inserted: inserted, kept: kept, skipped: skipped, batchId: batchId });
    } catch (e) {
      console.error('price-discrepancy upload error:', e.message);
      res.status(500).json({ error: e.message });
    }
  });

  // ── Manual single-row add (for someone without a file, or a one-off catch).
  // Superadmin/subadmin only — regular team members upload a file instead. ──
  app.post('/api/price-discrepancy/add', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var b = req.body || {};
      var customer = String(b.customerName || '').trim();
      var sku = String(b.skuCode || '').trim();
      var systemPrice = +b.systemPrice, lpoPrice = +b.lpoPrice;
      if (!customer || !sku) return res.status(400).json({ error: 'Customer Name and SKU Code are required.' });
      if (isNaN(systemPrice) || isNaN(lpoPrice)) return res.status(400).json({ error: 'System Price and LPO Price must be numbers.' });

      var existing = await findOpenMatch(customer, sku);
      if (existing) {
        await pool.query(
          `UPDATE price_discrepancies
           SET week_ending=$1, sku_description=COALESCE($2, sku_description), system_price=$3, lpo_price=$4,
               discrepancy=$5, remarks=COALESCE($6, remarks), uploaded_at=NOW(),
               reported_by_username=$7, reported_by_full_name=$8, times_reported = times_reported + 1
           WHERE id=$9`,
          [b.weekEnding || null, b.skuDescription || null, systemPrice, lpoPrice,
           Math.round((lpoPrice - systemPrice) * 100) / 100, b.remarks || null,
           req.user.username, req.user.full_name || req.user.username, existing.id]
        );
        auditLog(req.user.uid, req.user.username, 'PRICE_DISCREPANCY_ADD', 'Refreshed existing open entry ' + customer + ' / ' + sku, req.headers['x-forwarded-for'] || '');
        return res.json({ success: true, id: existing.id, kept: true });
      }

      var r = await pool.query(
        `INSERT INTO price_discrepancies
           (reported_by_username, reported_by_full_name, first_reported_by_username, first_reported_by_full_name,
            week_ending, customer_name, sku_code,
            sku_description, system_price, lpo_price, discrepancy, remarks, first_reported_at, times_reported)
         VALUES ($1,$2,$1,$2,$3,$4,$5,$6,$7,$8,$9,$10,NOW(),1) RETURNING id`,
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

  // ── Core query: every row, with is_repeated / days_unresolved computed
  // straight from its own stored first_reported_at + times_reported (no
  // grouping needed anymore — re-uploads refresh the existing row in place
  // instead of creating duplicates, so each row IS the current state of that
  // issue). This single query powers the consolidated view, the escalation
  // view, and both export endpoints, so the "repeated" definition never
  // drifts between them. ──
  async function fetchGrouped(whereSql, params) {
    var sql = `
      SELECT *,
        lower(regexp_replace(trim(customer_name), '\\s+', ' ', 'g')) AS cust_key,
        lower(regexp_replace(trim(sku_code), '[^a-zA-Z0-9]', '', 'g')) AS sku_key,
        (status = 'Open') AS still_open,
        (times_reported > 1 AND status = 'Open') AS is_repeated,
        CASE WHEN status = 'Open' THEN EXTRACT(DAY FROM (NOW() - first_reported_at))::int ELSE NULL END AS days_unresolved
      FROM price_discrepancies k
      ${whereSql || ''}
      ORDER BY days_unresolved DESC NULLS LAST, uploaded_at DESC`;
    return pool.query(sql, params || []);
  }

  // ── User-Wise Summary — per team member: how much they've logged, how much
  // is still open, how much of that is repeated, and when they last uploaded.
  // This is the "who is actually reporting vs who's gone quiet" view. ──
  async function fetchUserSummary() {
    var totals = await pool.query(
      `SELECT reported_by_username, reported_by_full_name,
              COUNT(*) AS total_logged,
              COUNT(*) FILTER (WHERE status = 'Open') AS open_count,
              COUNT(*) FILTER (WHERE status = 'Resolved') AS resolved_count,
              MAX(uploaded_at) AS last_submission
       FROM price_discrepancies
       GROUP BY reported_by_username, reported_by_full_name
       ORDER BY total_logged DESC`
    );
    // Repeated instances per user — how many of THEIR rows belong to a
    // Customer+SKU thread that is repeated and still open.
    var grouped = await fetchGrouped('', []);
    var repeatedByUser = {};
    grouped.rows.forEach(function (row) {
      if (row.is_repeated && row.still_open) {
        repeatedByUser[row.reported_by_username] = (repeatedByUser[row.reported_by_username] || 0) + 1;
      }
    });
    return totals.rows.map(function (row) {
      return Object.assign({}, row, { repeated_count: repeatedByUser[row.reported_by_username] || 0 });
    });
  }

  app.get('/api/price-discrepancy/user-summary', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var rows = await fetchUserSummary();
      res.json({ success: true, rows: rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── SKU Repeated Across Customers — same SKU, flagged by multiple DIFFERENT
  // customers. This is the systemic-pricing-issue view (master price is wrong
  // for that SKU) as opposed to the per-customer "Repeated" thread above. ──
  async function fetchSkuAcrossCustomers() {
    var r = await pool.query(
      `WITH keyed AS (
         SELECT *,
           lower(regexp_replace(trim(sku_code), '[^a-zA-Z0-9]', '', 'g')) AS sku_key
         FROM price_discrepancies
       )
       SELECT sku_key,
         (ARRAY_AGG(sku_code ORDER BY uploaded_at DESC))[1] AS sku_code,
         (ARRAY_AGG(sku_description ORDER BY uploaded_at DESC))[1] AS sku_description,
         COUNT(DISTINCT lower(trim(customer_name))) AS distinct_customers,
         COUNT(*) AS times_reported,
         COUNT(*) FILTER (WHERE status = 'Open') AS still_open_count,
         STRING_AGG(DISTINCT customer_name, '; ' ORDER BY customer_name) AS customer_list,
         (ARRAY_AGG(system_price ORDER BY uploaded_at DESC))[1] AS latest_system_price,
         (ARRAY_AGG(lpo_price ORDER BY uploaded_at DESC))[1] AS latest_lpo_price
       FROM keyed
       GROUP BY sku_key
       HAVING COUNT(DISTINCT lower(trim(customer_name))) > 1
       ORDER BY distinct_customers DESC, times_reported DESC`
    );
    return r.rows;
  }

  app.get('/api/price-discrepancy/sku-across-customers', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var rows = await fetchSkuAcrossCustomers();
      res.json({ success: true, rows: rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

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

  // ── Superadmin/subadmin: remove ONE wrong entry outright (not the whole
  // upload batch) and automatically warn whoever last touched it. The warning
  // lands as a real Task Tracker task assigned to that person, so it shows up
  // as a notification badge on their sidebar the same way any other task
  // does \u2014 no separate messaging system needed. Requires a reason so the
  // person removing it has to say why, and so the warning message isn't blank. ──
  app.delete('/api/price-discrepancy/:id/remove-and-warn', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var reason = String((req.body || {}).reason || '').trim();
      if (!reason) return res.status(400).json({ error: 'Please explain what was wrong with this entry \u2014 that explanation is what gets sent to the person who logged it.' });

      var row = await pool.query('SELECT * FROM price_discrepancies WHERE id = $1', [req.params.id]);
      if (!row.rows.length) return res.status(404).json({ error: 'Not found.' });
      var d = row.rows[0];
      var warnUsername = d.reported_by_username;
      var warnFullName = d.reported_by_full_name || d.reported_by_username;

      await pool.query('DELETE FROM price_discrepancies WHERE id = $1', [req.params.id]);

      // Auto-warning, delivered as a Task Tracker task \u2014 same table the Task
      // Tracker module already reads from, so it appears instantly with no
      // extra wiring. Deadline of today+1 so it shows as overdue quickly if ignored.
      var deadline = new Date(); deadline.setDate(deadline.getDate() + 1);
      var taskTitle = '\u26a0 Price discrepancy entry removed \u2014 please review';
      var taskComments = 'Your entry for ' + d.customer_name + ' / SKU ' + d.sku_code +
        ' (System: ' + d.system_price + ', LPO: ' + d.lpo_price + ') was removed by ' +
        (req.user.full_name || req.user.username) + '.\n\nReason: ' + reason +
        '\n\nPlease double-check this SKU/customer and re-upload if it\'s a genuine discrepancy, or drop it if it was a mistake.';
      try {
        await pool.query(
          `INSERT INTO tasks (title, comments, assigned_to_username, assigned_to_full_name, assigned_by_username, assigned_by_full_name, deadline)
           VALUES ($1,$2,$3,$4,$5,$6,$7)`,
          [taskTitle, taskComments, warnUsername, warnFullName, req.user.username, req.user.full_name || req.user.username, deadline.toISOString().split('T')[0]]
        );
      } catch (taskErr) {
        // Task Tracker table may not exist yet if that module isn't loaded \u2014 the
        // removal itself must still succeed even if the warning couldn't be created.
        console.error('Could not create warning task (removal still succeeded):', taskErr.message);
      }

      await auditLog(req.user.uid, req.user.username, 'PRICE_DISCREPANCY_REMOVE_AND_WARN',
        'Removed entry \u2014 ' + d.customer_name + ' / ' + d.sku_code + ' (was reported by ' + warnFullName + ') \u2014 reason: ' + reason,
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, warnedUser: warnFullName });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── A user's own submissions only ──
  app.get('/api/price-discrepancy/my-entries', requireAuth, async function (req, res) {
    try {
      var r = await fetchGrouped('WHERE k.reported_by_username = $1', [req.user.username]);
      res.json({ success: true, rows: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Shared export builder — detail rows + (optionally, for the full
  // consolidated download) User-Wise Summary and SKU-Across-Customers sheets,
  // plus the Executive Escalation sheet (repeated + still-open threads only). ──
  async function buildWorkbook(detailRows, scopeLabel, opts) {
    opts = opts || {};
    var ExcelJS = require('exceljs');
    var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', GOLD = 'FFFCF3CF', RED = 'FFFADBD8';
    function headerRow(ws, cells) {
      var row = ws.addRow(cells);
      row.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      return row;
    }

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    if (opts.userSummary) {
      var us = wb.addWorksheet('User-Wise Summary');
      us.columns = [{ width: 26 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 22 }];
      var ut = us.addRow(['USER-WISE SUMMARY \u2014 ALL TEAM MEMBERS']);
      us.mergeCells('A' + ut.number + ':F' + ut.number);
      ut.font = { bold: true, size: 13, color: { argb: NAVY } };
      us.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      us.addRow([]);
      headerRow(us, ['Team Member', 'Total Logged', 'Open', 'Resolved', 'Repeated (Open)', 'Last Submission']);
      opts.userSummary.forEach(function (row) {
        var r = us.addRow([
          row.reported_by_full_name || row.reported_by_username,
          +row.total_logged, +row.open_count, +row.resolved_count, +row.repeated_count,
          row.last_submission ? new Date(row.last_submission).toLocaleString('en-AE') : ''
        ]);
        if (+row.repeated_count > 0) r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } }; });
      });
    }

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
      // System Price higher than LPO Price (discrepancy negative) -> red. System Price
      // lower than LPO Price (discrepancy positive) -> green. Takes priority over the
      // repeated-row highlight above, just on these 3 price cells.
      var variance = +row.discrepancy;
      var priceBg = variance < 0 ? 'FFF5C6C0' : 'FFC8E6C9';
      var priceTxt = variance < 0 ? 'FF7A241C' : 'FF1B5E20';
      [6, 7, 8].forEach(function (col) {
        var c = r.getCell(col);
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: priceBg } };
        c.font = { color: { argb: priceTxt }, bold: col === 8 };
      });
    });

    if (opts.skuAcrossCustomers) {
      var sc = wb.addWorksheet('SKU Across Customers');
      sc.columns = [{ width: 16 }, { width: 34 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 60 }];
      var sct = sc.addRow(['SAME SKU \u2014 REPORTED BY MULTIPLE CUSTOMERS']);
      sc.mergeCells('A' + sct.number + ':H' + sct.number);
      sct.font = { bold: true, size: 13, color: { argb: NAVY } };
      sc.addRow(['Likely a systemic/master-price issue for that SKU, not a single customer\u2019s pricing problem.']);
      sc.addRow([]);
      headerRow(sc, ['SKU Code', 'Description', 'Distinct Customers', 'Times Reported', 'Still Open', 'Latest System Price', 'Latest LPO Price', 'Customers']);
      (opts.skuAcrossCustomers || []).forEach(function (row) {
        var r = sc.addRow([
          row.sku_code, row.sku_description || '', +row.distinct_customers, +row.times_reported,
          +row.still_open_count, +row.latest_system_price, +row.latest_lpo_price, row.customer_list
        ]);
        if (+row.distinct_customers >= 3) r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } }; });
        else r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
      });
      if (!(opts.skuAcrossCustomers || []).length) sc.addRow(['No SKU currently flagged by more than one customer.']);
    }

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
      // The User-Wise Summary and SKU-Across-Customers sheets are always the
      // full picture (not scoped to a single rep filter) since they only make
      // sense compared across the whole team.
      var userSummary = await fetchUserSummary();
      var skuAcrossCustomers = await fetchSkuAcrossCustomers();
      var wb = await buildWorkbook(r.rows, label, { userSummary: userSummary, skuAcrossCustomers: skuAcrossCustomers });
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

  // ── Sales & Marketing Summary — a rollup, not a raw dump, meant to be sent
  // straight to Sales/Marketing so they know which customers and which brands
  // currently have live pricing issues to go fix. Brand is inferred from the
  // SKU prefix before the first hyphen (e.g. "STM-312-0" -> brand "STM").
  // Defaults to OPEN issues only, since resolved ones aren't actionable. ──
  function skuBrand(skuCode) {
    var s = String(skuCode || '').trim().toUpperCase();
    var m = s.match(/^([A-Z]+)/);
    return m ? m[1] : (s.split('-')[0] || 'UNKNOWN');
  }
  app.get('/api/price-discrepancy/summary/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var includeResolved = req.query.include_resolved === 'true';
      var where = includeResolved ? '' : `WHERE k.status = 'Open'`;
      var r = await fetchGrouped(where, []);
      var rows = r.rows;

      var ExcelJS = require('exceljs');
      var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', GOLD = 'FFFCF3CF';
      var RED_BG = 'FFF5C6C0', RED_TEXT = 'FF7A241C', GREEN_BG = 'FFC8E6C9', GREEN_TEXT = 'FF1B5E20';
      function headerRow(ws, cells) {
        var row = ws.addRow(cells);
        row.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
        return row;
      }
      // System Price higher, LPO Price lower (variance = LPO - System is negative) -> RED.
      // System Price lower, LPO Price higher (variance positive) -> GREEN.
      // Colors System Price, LPO Price, and Discrepancy together (columns 5,6,7 in these
      // two sheets) so the direction is obvious without reading the numbers.
      function applyPriceColors(row, variance) {
        var bg = variance < 0 ? RED_BG : GREEN_BG;
        var txt = variance < 0 ? RED_TEXT : GREEN_TEXT;
        [5, 6, 7].forEach(function (col) {
          var c = row.getCell(col);
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
          c.font = { color: { argb: txt }, bold: col === 7 };
        });
      }
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();

      // ── Sheet 1: Brand & Customer — every discrepancy on its own line, sorted
      // Brand -> Customer -> SKU. Alternating shade per Brand block so the eye
      // can tell where one brand ends and the next starts, without merging
      // anything into a sum. ──
      var detailRows = rows.map(function (row) {
        return {
          brand: skuBrand(row.sku_code), customer: row.customer_name, sku: row.sku_code,
          desc: row.sku_description || '', system: +row.system_price, lpo: +row.lpo_price,
          variance: +row.discrepancy, days: row.days_unresolved
        };
      });

      // ── Sheet 0: Executive Summary — one line per Brand: how many DISTINCT
      // SKUs currently have a price discrepancy, and what share of every
      // affected SKU across the whole business that brand represents. This is
      // the "read this first" sheet before the two detail sheets below. ──
      var brandSkuSets = {}, brandCustSets = {}, brandRowCount = {}, brandLowerSkus = {}, brandHigherSkus = {};
      detailRows.forEach(function (d) {
        if (!brandSkuSets[d.brand]) { brandSkuSets[d.brand] = new Set(); brandCustSets[d.brand] = new Set(); brandRowCount[d.brand] = 0; brandLowerSkus[d.brand] = new Set(); brandHigherSkus[d.brand] = new Set(); }
        brandSkuSets[d.brand].add(d.sku);
        brandCustSets[d.brand].add(d.customer);
        brandRowCount[d.brand]++;
        // variance = LPO - System. Positive -> System Price is LOWER than LPO.
        // Negative -> System Price is HIGHER than LPO.
        if (d.variance > 0) brandLowerSkus[d.brand].add(d.sku); else if (d.variance < 0) brandHigherSkus[d.brand].add(d.sku);
      });
      var totalDistinctSkus = new Set(detailRows.map(function (d) { return d.sku; })).size;
      var execSummary = Object.keys(brandSkuSets).map(function (brand) {
        var skuCount = brandSkuSets[brand].size;
        return {
          brand: brand,
          skuCount: skuCount,
          pctOfTotal: totalDistinctSkus > 0 ? +((skuCount / totalDistinctSkus) * 100).toFixed(1) : 0,
          customerCount: brandCustSets[brand].size,
          occurrences: brandRowCount[brand],
          systemLowerSkus: brandLowerSkus[brand].size,
          systemHigherSkus: brandHigherSkus[brand].size
        };
      }).sort(function (a, b) { return b.skuCount - a.skuCount; });

      // ── The "where do we focus" headline: of every SKU currently flagged,
      // what share has System Price sitting BELOW the LPO price (we're likely
      // undercharging / leaking margin) vs ABOVE it (LPO likely underpriced
      // vs system, a customer-facing dispute risk). One line, whole business. ──
      var allSkusLower = new Set(), allSkusHigher = new Set();
      detailRows.forEach(function (d) {
        if (d.variance > 0) allSkusLower.add(d.sku); else if (d.variance < 0) allSkusHigher.add(d.sku);
      });
      var lowerPct = totalDistinctSkus > 0 ? +((allSkusLower.size / totalDistinctSkus) * 100).toFixed(1) : 0;
      var higherPct = totalDistinctSkus > 0 ? +((allSkusHigher.size / totalDistinctSkus) * 100).toFixed(1) : 0;

      var ex = wb.addWorksheet('Executive Summary');
      ex.columns = [{ width: 14 }, { width: 20 }, { width: 12 }, { width: 20 }, { width: 24 }, { width: 20 }, { width: 20 }];
      var exTitle = ex.addRow(['EXECUTIVE SUMMARY \u2014 SKUs WITH PRICE DISCREPANCY, BY BRAND' + (includeResolved ? '' : ' (OPEN ISSUES ONLY)')]);
      ex.mergeCells('A' + exTitle.number + ':G' + exTitle.number);
      exTitle.font = { bold: true, size: 13, color: { argb: NAVY } };
      ex.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      ex.addRow(['Total distinct SKUs currently affected (all brands)', totalDistinctSkus]);
      ex.addRow([]);

      // Where to focus — one line, whole business.
      var focusTitle = ex.addRow(['\u25b6 WHERE TO FOCUS']);
      focusTitle.font = { bold: true, size: 12, color: { argb: NAVY } };
      var focusRow = ex.addRow(['System Price LOWER than LPO', allSkusLower.size + ' SKUs (' + lowerPct + '%)', '', 'System Price HIGHER than LPO', allSkusHigher.size + ' SKUs (' + higherPct + '%)']);
      focusRow.getCell(1).font = { bold: true };
      focusRow.getCell(2).font = { bold: true, color: { argb: 'FF1B5E20' } };
      focusRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8E6C9' } };
      focusRow.getCell(4).font = { bold: true };
      focusRow.getCell(5).font = { bold: true, color: { argb: 'FF7A241C' } };
      focusRow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5C6C0' } };
      ex.addRow(['   (we may be undercharging \u2014 margin leak)', '', '', '   (LPO looks underpriced vs system \u2014 customer dispute risk)']).font = { italic: true, size: 10, color: { argb: 'FF7A7A7A' } };
      ex.addRow([]);

      var brandTableTitle = ex.addRow(['\u25b6 BY BRAND']);
      brandTableTitle.font = { bold: true, size: 12, color: { argb: NAVY } };
      ex.addRow(['"SKUs Affected" = distinct SKUs with an open issue. "Occurrences" = every time it was logged \u2014 the same SKU reported by 3 different customers counts as 3 occurrences.']).font = { italic: true, size: 10, color: { argb: 'FF7A7A7A' } };
      headerRow(ex, ['Brand', 'SKUs Affected', '% of Total', 'Customers Affected', 'Occurrences (Times Logged)', 'System Lower (SKUs)', 'System Higher (SKUs)']);
      execSummary.forEach(function (g) {
        var row = ex.addRow([g.brand, g.skuCount, g.pctOfTotal + '%', g.customerCount, g.occurrences, g.systemLowerSkus, g.systemHigherSkus]);
        row.getCell(2).font = { bold: true };
        row.getCell(3).font = { color: { argb: g.pctOfTotal >= 15 ? 'FFB03A2E' : 'FF3B3B3B' }, bold: g.pctOfTotal >= 15 };
        row.getCell(6).font = { color: { argb: 'FF1B5E20' } };
        row.getCell(7).font = { color: { argb: 'FF7A241C' } };
      });
      if (!execSummary.length) ex.addRow(['No open price discrepancies right now.']);
      var totalRow = ex.addRow(['TOTAL', totalDistinctSkus, '100%',
        new Set(detailRows.map(function (d) { return d.customer; })).size,
        detailRows.length, allSkusLower.size, allSkusHigher.size]);
      totalRow.font = { bold: true };
      totalRow.eachCell(function (c) { c.border = { top: { style: 'thin' } }; });

      var byBrand = detailRows.slice().sort(function (a, b) {
        return a.brand !== b.brand ? a.brand.localeCompare(b.brand)
             : a.customer !== b.customer ? a.customer.localeCompare(b.customer)
             : a.sku.localeCompare(b.sku);
      });
      var bc = wb.addWorksheet('Brand & Customer Summary');
      bc.columns = [{ width: 14 }, { width: 34 }, { width: 16 }, { width: 30 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }];
      var bcTitle = bc.addRow(['BRAND & CUSTOMER PRICE DISCREPANCY' + (includeResolved ? '' : ' (OPEN ISSUES ONLY)') + ' \u2014 EVERY ITEM, GROUPED BY BRAND']);
      bc.mergeCells('A' + bcTitle.number + ':H' + bcTitle.number);
      bcTitle.font = { bold: true, size: 13, color: { argb: NAVY } };
      bc.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      bc.addRow([]);
      headerRow(bc, ['Brand', 'Customer', 'SKU Code', 'Description', 'System Price', 'LPO Price', 'Discrepancy (AED)', 'Days Unresolved']);
      var brandShadeOn = false, lastBrandSeen = null;
      byBrand.forEach(function (d) {
        if (lastBrandSeen !== null && d.brand !== lastBrandSeen) brandShadeOn = !brandShadeOn;
        lastBrandSeen = d.brand;
        var row = bc.addRow([d.brand, d.customer, d.sku, d.desc, d.system, d.lpo, d.variance, d.days == null ? '' : d.days]);
        if (brandShadeOn) row.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
        applyPriceColors(row, d.variance);
      });
      if (!byBrand.length) bc.addRow(['No open price discrepancies right now.']);

      // ── Sheet 2: Customer-Wise — same real rows, sorted Customer -> Brand -> SKU,
      // shaded per Customer block instead of per Brand. ──
      var byCustomer = detailRows.slice().sort(function (a, b) {
        return a.customer !== b.customer ? a.customer.localeCompare(b.customer)
             : a.brand !== b.brand ? a.brand.localeCompare(b.brand)
             : a.sku.localeCompare(b.sku);
      });
      var cw = wb.addWorksheet('Customer-Wise Discrepancy');
      cw.columns = [{ width: 34 }, { width: 14 }, { width: 16 }, { width: 30 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }];
      var cwTitle = cw.addRow(['CUSTOMER-WISE PRICE DISCREPANCY' + (includeResolved ? '' : ' (OPEN ISSUES ONLY)') + ' \u2014 EVERY ITEM, GROUPED BY CUSTOMER']);
      cw.mergeCells('A' + cwTitle.number + ':H' + cwTitle.number);
      cwTitle.font = { bold: true, size: 13, color: { argb: NAVY } };
      cw.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      cw.addRow([]);
      headerRow(cw, ['Customer', 'Brand', 'SKU Code', 'Description', 'System Price', 'LPO Price', 'Discrepancy (AED)', 'Days Unresolved']);
      var custShadeOn = false, lastCustSeen = null;
      byCustomer.forEach(function (d) {
        if (lastCustSeen !== null && d.customer !== lastCustSeen) custShadeOn = !custShadeOn;
        lastCustSeen = d.customer;
        var row = cw.addRow([d.customer, d.brand, d.sku, d.desc, d.system, d.lpo, d.variance, d.days == null ? '' : d.days]);
        if (custShadeOn) row.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
        applyPriceColors(row, d.variance);
      });
      if (!byCustomer.length) cw.addRow(['No open price discrepancies right now.']);

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Price_Discrepancy_Sales_Summary_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('price-discrepancy summary export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });
};
