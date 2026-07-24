// ══ CS WORKLOAD TRACKER ══════════════════════════════════
// Each CS rep uploads their own daily file of orders they personally processed
// (or, for GRV-only staff like Renan, GRV references generated). Since order
// processing isn't logged with a "handled by" field anywhere upstream, this
// self-reported log is the only source of truth for who did what — which is
// exactly why cross-checking for duplicate claims on the same order number
// matters: if two people log the same order, that's either double-work,
// confusion about ownership, or something worth a direct conversation.
const XLSX = require('xlsx');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog) {

  (async function initTables() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS cs_workload_entries (
        id SERIAL PRIMARY KEY,
        rep_username TEXT NOT NULL,
        rep_full_name TEXT,
        entry_type TEXT NOT NULL DEFAULT 'order', -- 'order' or 'grv'
        entry_date DATE,
        order_number TEXT,
        customer_name TEXT,
        resolution_type TEXT,   -- Price / UOM / Other — for entry_type='order'
        grv_reference TEXT,     -- for entry_type='grv'
        file_name TEXT,
        uploaded_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_cswl_order_number ON cs_workload_entries (order_number)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_cswl_rep ON cs_workload_entries (rep_username)`);
      // Historical baseline — the 3-month account-name-matched estimate. Kept in a
      // SEPARATE table from real daily uploads on purpose: it's a rough estimate,
      // not verified per-user data, and must never be blended with or mistaken for
      // the real, verified daily entries above.
      await pool.query(`CREATE TABLE IF NOT EXISTS cs_workload_baseline (
        id SERIAL PRIMARY KEY,
        person_name TEXT NOT NULL,
        account_label TEXT,
        total_orders INT DEFAULT 0,
        auto_orders INT DEFAULT 0,
        manual_orders INT DEFAULT 0,
        period_label TEXT,
        uploaded_by TEXT,
        uploaded_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      console.log('CS Workload module DB ready');
    } catch (e) { console.error('CS Workload init error:', e.message); }
  })();

  function normalizeHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
  function fc(keys, names) {
    return keys.find(function (k) {
      var nk = normalizeHeader(k);
      return names.some(function (n) { return nk.indexOf(n) !== -1; });
    }) || null;
  }

  // ── Upload: each rep uploads their own file. Identity comes from the logged-in
  // user (req.user), not from anything in the file itself — so nobody can upload
  // and claim to be someone else.
  app.post('/api/cs-workload/upload', requireAuth, upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var entryType = (req.body.entryType === 'grv') ? 'grv' : 'order';
      var wb;
      try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
      catch (e) { return res.status(400).json({ error: 'Could not read that file. Upload an .xlsx/.csv.' }); }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      var raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!raw.length) return res.status(400).json({ error: 'No rows found in that file.' });

      var keys = Object.keys(raw[0]);
      var C = {
        date: fc(keys, ['date']),
        // Prefer an exact "order number" match before falling back to anything
        // that merely contains "order" (avoids grabbing unrelated order-type columns).
        order: fc(keys, ['ordernumber']) || fc(keys, ['order']),
        // Prefer "customer name" specifically before falling back to any column
        // that merely contains "customer" (avoids grabbing an LPO/reference column
        // that happens to have "customer" in its header too).
        customer: fc(keys, ['customername']) || fc(keys, ['customer']),
        resolution: fc(keys, ['resolution', 'reason']),
        grv: fc(keys, ['grvreference', 'grv'])
      };
      if (!C.order) return res.status(400).json({ error: 'Could not find an Order Number column in that file.' });

      var repUsername = req.user.username;
      var repFullName = req.user.full_name || req.user.username;
      var inserted = 0;

      for (var i = 0; i < raw.length; i++) {
        var row = raw[i];
        var orderNum = String(row[C.order] || '').trim();
        if (!orderNum) continue;
        var entryDate = C.date ? row[C.date] : null;
        var customer = C.customer ? String(row[C.customer] || '').trim() : null;
        var resolution = C.resolution ? String(row[C.resolution] || '').trim() : null;
        var grv = C.grv ? String(row[C.grv] || '').trim() : null;
        await pool.query(
          `INSERT INTO cs_workload_entries (rep_username, rep_full_name, entry_type, entry_date, order_number, customer_name, resolution_type, grv_reference, file_name)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
          [repUsername, repFullName, entryType, entryDate, orderNum, customer, resolution, grv, req.file.originalname]
        );
        inserted++;
      }

      auditLog(req.user.uid, repUsername, 'CS_WORKLOAD_UPLOAD', 'Uploaded ' + inserted + ' ' + entryType + ' entries (' + req.file.originalname + ')', req.headers['x-forwarded-for'] || '');
      res.json({ success: true, inserted: inserted });
    } catch (e) {
      console.error('cs-workload upload error:', e.message);
      res.status(500).json({ error: e.message });
    }
  });

  // ── Scorecard: per-rep totals, optionally filtered by date range ──
  // MANAGEMENT ONLY — regular reps get their own view via /my-summary instead,
  // so they can't see everyone else's numbers.
  app.get('/api/cs-workload/scorecard', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var r = await pool.query(
        `SELECT rep_username, rep_full_name, entry_type, COUNT(*) AS cnt, COUNT(DISTINCT customer_name) AS distinct_customers
         FROM cs_workload_entries ${where}
         GROUP BY rep_username, rep_full_name, entry_type
         ORDER BY cnt DESC`, params
      );
      res.json({ success: true, rows: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Personal view: a regular rep sees ONLY their own uploads — their daily
  // counts (for a chart) and a resolution-type percentage breakdown. No visibility
  // into anyone else's numbers, and no team comparison, by design. ──
  app.get('/api/cs-workload/my-summary', requireAuth, async function (req, res) {
    try {
      var username = req.user.username;
      var daily = await pool.query(
        `SELECT entry_date, entry_type, COUNT(*) AS cnt FROM cs_workload_entries
         WHERE rep_username = $1 GROUP BY entry_date, entry_type ORDER BY entry_date`, [username]
      );
      var breakdown = await pool.query(
        `SELECT COALESCE(NULLIF(resolution_type,''),'Unspecified') AS resolution_type, COUNT(*) AS cnt
         FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order' GROUP BY resolution_type`, [username]
      );
      var totals = await pool.query(
        `SELECT entry_type, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1 GROUP BY entry_type`, [username]
      );
      res.json({ success: true, daily: daily.rows, breakdown: breakdown.rows, totals: totals.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Historical baseline: 3-month account-matched estimate. Management uploads
  // it once (replaces whatever was there before, like other snapshot uploads in
  // this app), everyone can view/download it, clearly labeled as an estimate. ──
  app.post('/api/cs-workload/baseline/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try { wb = XLSX.read(req.file.buffer, { type: 'buffer' }); }
      catch (e) { return res.status(400).json({ error: 'Could not read that file.' }); }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      // Read as raw rows first (no assumed header) — reports like this one have a
      // title, generated-date, and TEAM TOTALS block before the real per-person
      // table, so row 1 is almost never the actual header.
      var grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!grid.length) return res.status(400).json({ error: 'No rows found in that file.' });

      function normRow(row) { return row.map(function (c) { return normalizeHeader(c); }); }
      var headerRowIdx = -1, C = null;
      for (var i = 0; i < grid.length; i++) {
        var nr = normRow(grid[i]);
        var pIdx = nr.findIndex(function (c) { return c.indexOf('person') !== -1 || c.indexOf('rep') !== -1; });
        var mIdx = nr.findIndex(function (c) { return c.indexOf('manual') !== -1; });
        if (pIdx !== -1 && mIdx !== -1) {
          headerRowIdx = i;
          C = {
            person: pIdx,
            account: nr.findIndex(function (c) { return c.indexOf('account') !== -1; }),
            total: nr.findIndex(function (c) { return c.indexOf('total') !== -1; }),
            auto: nr.findIndex(function (c) { return c.indexOf('auto') !== -1; }),
            manual: mIdx
          };
          break;
        }
      }
      if (headerRowIdx === -1) return res.status(400).json({ error: 'Could not find a header row with Person and Manual columns in that file.' });

      var periodLabel = req.body.periodLabel || '1 Apr \u2013 20 Jul 2026 (estimated)';
      // Collect rows below the header, stopping at the first fully-blank row
      // (that's the end of the real table — everything after is notes/footers).
      var merged = {}; // person -> {account, total, auto, manual} — sums duplicates
      for (var r = headerRowIdx + 1; r < grid.length; r++) {
        var row = grid[r];
        var isBlank = row.every(function (c) { return String(c).trim() === ''; });
        if (isBlank) break;
        var person = String(row[C.person] || '').trim();
        if (!person) continue;
        var total = C.total !== -1 ? (+row[C.total] || 0) : 0;
        var auto = C.auto !== -1 ? (+row[C.auto] || 0) : 0;
        var manual = C.manual !== -1 ? (+row[C.manual] || 0) : 0;
        var account = C.account !== -1 ? String(row[C.account] || '').trim() : '';
        if (!merged[person]) merged[person] = { account: account, total: 0, auto: 0, manual: 0 };
        merged[person].total += total;
        merged[person].auto += auto;
        merged[person].manual += manual;
        if (account && !merged[person].account) merged[person].account = account;
      }
      var personNames = Object.keys(merged);
      if (!personNames.length) return res.status(400).json({ error: 'Found a header row but no data rows underneath it.' });

      await pool.query('DELETE FROM cs_workload_baseline');
      var inserted = 0;
      for (var p = 0; p < personNames.length; p++) {
        var name = personNames[p];
        var v = merged[name];
        await pool.query(
          `INSERT INTO cs_workload_baseline (person_name, account_label, total_orders, auto_orders, manual_orders, period_label, uploaded_by)
           VALUES ($1,$2,$3,$4,$5,$6,$7)`,
          [name, v.account || null, v.total, v.auto, v.manual, periodLabel, req.user.username]
        );
        inserted++;
      }
      auditLog(req.user.uid, req.user.username, 'CS_WORKLOAD_BASELINE_UPLOAD', 'Replaced baseline with ' + inserted + ' people (' + personNames.join(', ') + ')', req.headers['x-forwarded-for'] || '');
      res.json({ success: true, inserted: inserted });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/cs-workload/baseline', requireAuth, async function (req, res) {
    try {
      var r = await pool.query('SELECT * FROM cs_workload_baseline ORDER BY manual_orders DESC');
      res.json({ success: true, rows: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/cs-workload/baseline/export', requireAuth, async function (req, res) {
    try {
      var ExcelJS = require('exceljs');
      var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', GOLD = 'FFFCF3CF';
      var r = await pool.query('SELECT * FROM cs_workload_baseline ORDER BY manual_orders DESC');
      var teamManualTotal = r.rows.reduce(function (s, row) { return s + (+row.manual_orders); }, 0);
      var wb = new ExcelJS.Workbook();
      var ws = wb.addWorksheet('3-Month Baseline (Estimated)');
      ws.columns = [{ width: 26 }, { width: 34 }, { width: 12 }, { width: 12 }, { width: 10 }, { width: 12 }, { width: 10 }, { width: 16 }];
      var t = ws.addRow(['3-MONTH HISTORICAL BASELINE \u2014 ESTIMATED FROM ACCOUNT-NAME MATCHING, NOT VERIFIED PER-USER DATA']);
      ws.mergeCells('A' + t.number + ':H' + t.number);
      t.font = { bold: true, size: 12, color: { argb: NAVY } };
      ws.addRow([]);
      var hdr = ws.addRow(['Person', 'Account', 'Total', 'Auto', 'Auto %', 'Manual', 'Manual %', 'Share of Team Manual']);
      hdr.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      r.rows.forEach(function (row) {
        var total = +row.total_orders, auto = +row.auto_orders, manual = +row.manual_orders;
        var autoPct = total ? Math.round(auto / total * 1000) / 10 : 0;
        var manualPct = total ? Math.round(manual / total * 1000) / 10 : 0;
        var teamShare = teamManualTotal ? Math.round(manual / teamManualTotal * 1000) / 10 : 0;
        var rr = ws.addRow([row.person_name, row.account_label, total, auto, autoPct + '%', manual, manualPct + '%', teamShare + '%']);
        rr.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
      });
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="CS_Workload_Baseline_Estimated.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
  });

  // ── Duplicate order claims: same order number logged by more than one rep ──
  // MANAGEMENT ONLY, same reasoning as the scorecard above.
  app.get('/api/cs-workload/duplicates', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var dupOrders = await pool.query(`
        SELECT order_number
        FROM cs_workload_entries
        WHERE entry_type = 'order' AND order_number IS NOT NULL AND order_number != ''
        GROUP BY order_number
        HAVING COUNT(DISTINCT rep_username) > 1
      `);
      if (!dupOrders.rows.length) return res.json({ success: true, duplicates: [] });
      var orderNums = dupOrders.rows.map(function (r) { return r.order_number; });
      var detail = await pool.query(
        `SELECT order_number, rep_username, rep_full_name, entry_date, customer_name, resolution_type, uploaded_at
         FROM cs_workload_entries
         WHERE entry_type = 'order' AND order_number = ANY($1)
         ORDER BY order_number, uploaded_at`, [orderNums]
      );
      var grouped = {};
      detail.rows.forEach(function (row) {
        if (!grouped[row.order_number]) grouped[row.order_number] = [];
        grouped[row.order_number].push(row);
      });
      var duplicates = Object.keys(grouped).map(function (k) { return { order_number: k, claims: grouped[k] }; });
      res.json({ success: true, duplicates: duplicates });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Monthly color export: Scorecard + Duplicate Claims ──
  app.get('/api/cs-workload/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var ExcelJS = require('exceljs');
      var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', RED = 'FFFDE0DE', GOLD = 'FFFCF3CF', GREEN = 'FFE1F5E9';
      function headerRow(ws, cells) {
        var row = ws.addRow(cells);
        row.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
        return row;
      }
      var params = [], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var scoreR = await pool.query(
        `SELECT rep_username, rep_full_name, entry_type, COUNT(*) AS cnt, COUNT(DISTINCT customer_name) AS distinct_customers
         FROM cs_workload_entries ${where} GROUP BY rep_username, rep_full_name, entry_type ORDER BY cnt DESC`, params
      );
      var dupOrders = await pool.query(`
        SELECT order_number FROM cs_workload_entries
        WHERE entry_type='order' AND order_number IS NOT NULL AND order_number != ''
        GROUP BY order_number HAVING COUNT(DISTINCT rep_username) > 1
      `);
      var dupDetail = { rows: [] };
      if (dupOrders.rows.length) {
        var orderNums = dupOrders.rows.map(function (r) { return r.order_number; });
        dupDetail = await pool.query(
          `SELECT order_number, rep_full_name, entry_date, customer_name, resolution_type FROM cs_workload_entries
           WHERE entry_type='order' AND order_number = ANY($1) ORDER BY order_number`, [orderNums]
        );
      }

      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();

      var sc = wb.addWorksheet('Scorecard');
      sc.columns = [{ width: 22 }, { width: 14 }, { width: 12 }, { width: 18 }];
      var t = sc.addRow(['CS WORKLOAD SCORECARD']);
      sc.mergeCells('A' + t.number + ':D' + t.number);
      t.font = { bold: true, size: 13, color: { argb: NAVY } };
      sc.addRow(['Period', (req.query.date_from || 'All time') + ' to ' + (req.query.date_to || 'Latest')]);
      sc.addRow([]);
      headerRow(sc, ['Rep', 'Entry Type', 'Count', 'Distinct Customers']);
      scoreR.rows.forEach(function (row) {
        sc.addRow([row.rep_full_name || row.rep_username, row.entry_type, +row.cnt, +row.distinct_customers]);
      });

      var dup = wb.addWorksheet('Duplicate Order Claims');
      dup.columns = [{ width: 18 }, { width: 22 }, { width: 14 }, { width: 28 }, { width: 16 }];
      var t2 = dup.addRow(['DUPLICATE ORDER CLAIMS \u2014 SAME ORDER LOGGED BY MORE THAN ONE PERSON']);
      dup.mergeCells('A' + t2.number + ':E' + t2.number);
      t2.font = { bold: true, size: 12, color: { argb: NAVY } };
      dup.addRow([]);
      headerRow(dup, ['Order Number', 'Rep', 'Date', 'Customer', 'Resolution']);
      var lastOrder = null, colorToggle = false;
      dupDetail.rows.forEach(function (row) {
        if (row.order_number !== lastOrder) { colorToggle = !colorToggle; lastOrder = row.order_number; }
        var r = dup.addRow([row.order_number, row.rep_full_name, row.entry_date, row.customer_name, row.resolution_type]);
        r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorToggle ? GOLD : RED } }; });
      });
      if (!dupDetail.rows.length) {
        var okRow = dup.addRow(['No duplicate order claims found in this period.']);
        okRow.font = { color: { argb: 'FF2E7D4F' }, bold: true };
      }

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="CS_Workload_Scorecard_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('cs-workload export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });

  app.delete('/api/cs-workload/clear', requireAuth, requireRole('superadmin'), async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM cs_workload_entries');
      await auditLog(req.user.uid, req.user.username, 'CS_WORKLOAD_CLEAR', 'Cleared all CS workload entries (' + r.rowCount + ')', req.headers['x-forwarded-for'] || '');
      res.json({ success: true, deleted: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
};
