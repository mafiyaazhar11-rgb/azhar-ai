// ══ CS WORKLOAD TRACKER ══════════════════════════════════
// Each CS rep uploads their own daily file of orders they personally processed
// (or, for GRV-only staff like Renan, GRV references generated). Since order
// processing isn't logged with a "handled by" field anywhere upstream, this
// self-reported log is the only source of truth for who did what — which is
// exactly why cross-checking for duplicate claims on the same order number
// matters: if two people log the same order, that's either double-work,
// confusion about ownership, or something worth a direct conversation.
const XLSX = require('xlsx');
const crypto = require('crypto');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, anthropic) {

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
        booking_status TEXT DEFAULT 'Manual', -- 'Auto' or 'Manual' — real per-row status if the
                                               -- uploaded file has it, otherwise defaults to Manual
                                               -- (since historically only manually-touched orders got logged)
        fob_status TEXT,        -- raw 3-category text: AUTO BOOKED / AUTO BOOKED W/ MANUAL
                                 -- INTERVENTION / MANUAL BOOKED — kept alongside the binary
                                 -- booking_status above so the full breakdown isn't lost
        grv_reference TEXT,     -- for entry_type='grv'
        file_name TEXT,
        uploaded_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`ALTER TABLE cs_workload_entries ADD COLUMN IF NOT EXISTS booking_status TEXT DEFAULT 'Manual'`);
      await pool.query(`ALTER TABLE cs_workload_entries ADD COLUMN IF NOT EXISTS fob_status TEXT`);
      // batch_id: one value per upload request (not per file_name/uploaded_at, which can
      // collide across re-uploads) — this is what makes a single wrong upload deletable
      // on its own, without touching anyone else's data or requiring a full Clear All.
      await pool.query(`ALTER TABLE cs_workload_entries ADD COLUMN IF NOT EXISTS batch_id TEXT`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_cswl_order_number ON cs_workload_entries (order_number)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_cswl_rep ON cs_workload_entries (rep_username)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_cswl_batch ON cs_workload_entries (batch_id)`);
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
        // Prioritize "Booked Date" specifically — this file also has "Date Ordered" and
        // "Request Date", both of which also contain "date" and could be picked up wrongly.
        // Booked Date is what actually matters: when the order was booked, not when it
        // was originally requested.
        date: fc(keys, ['bookeddate']) || fc(keys, ['date']),
        // Prefer an exact "order number" match before falling back to anything
        // that merely contains "order" (avoids grabbing unrelated order-type columns).
        order: fc(keys, ['ordernumber']) || fc(keys, ['order']),
        // Prefer "customer name" specifically before falling back to any column
        // that merely contains "customer" (avoids grabbing an LPO/reference column
        // that happens to have "customer" in its header too).
        customer: fc(keys, ['customername']) || fc(keys, ['customer']),
        // "Trading Partner Information" (NO ISSUE / PRICE ISSUE / CASH ORDER / INP) is
        // the reason text — but NOT the authoritative Auto/Manual signal (see status below).
        resolution: fc(keys, ['tradingpartner']) || fc(keys, ['resolution', 'reason']),
        grv: fc(keys, ['grvreference', 'grv']),
        // FOB is the real, authoritative Auto/Manual field in this file format:
        // "AUTO BOOKED" / "AUTO BOOKED W/ MANUAL INTERVENTION" / "MANUAL BOOKED".
        // Trading Partner Info can look consistent with it but FOB is the source of truth.
        status: fc(keys, ['fob']) || fc(keys, ['bookingstatus', 'autobookingstatus', 'status'])
      };
      if (!C.order) return res.status(400).json({ error: 'Could not find an Order Number column in that file.' });

      var repUsername = req.user.username;
      var repFullName = req.user.full_name || req.user.username;
      var batchId = crypto.randomUUID();
      var inserted = 0, updated = 0;

      for (var i = 0; i < raw.length; i++) {
        var row = raw[i];
        var orderNum = String(row[C.order] || '').trim();
        if (!orderNum) continue;
        var entryDate = C.date ? row[C.date] : null;
        var customer = C.customer ? String(row[C.customer] || '').trim() : null;
        var grv = C.grv ? String(row[C.grv] || '').trim() : null;
        var rawReason = C.resolution ? String(row[C.resolution] || '').trim() : '';
        var isNoIssue = /no\s*issue/i.test(rawReason);
        var isIgnorable = isNoIssue || /^inp$/i.test(rawReason); // INP = internal comment, not a real reason category
        // "NO ISSUE" or "INP" = no meaningful reason to log. Anything else (PRICE ISSUE,
        // CASH ORDER, etc.) keeps the exact wording from the file as the resolution reason.
        var resolution = (rawReason && !isIgnorable) ? rawReason : null;
        var bookingStatus = 'Manual';
        var fobStatus = null;
        if (C.status) {
          // FOB is authoritative: check for "manual" FIRST, since "AUTO BOOKED W/ MANUAL
          // INTERVENTION" contains both words — manual intervention happened, so it counts
          // as real work done, even though the order was technically auto-booked initially.
          var rawStatusRaw = String(row[C.status] || '').trim();
          fobStatus = rawStatusRaw || null;
          var rawStatus = rawStatusRaw.toLowerCase();
          bookingStatus = rawStatus.indexOf('manual') !== -1 ? 'Manual' : (rawStatus.indexOf('auto') !== -1 ? 'Auto' : 'Manual');
        }

        // Upload never blocks. Two different things can happen with a repeated order
        // number, and they're handled differently on purpose:
        //  - SAME user re-logging the SAME order (even on a different day/upload) is the
        //    real inflation risk ("I have more orders") — so that just UPDATES their
        //    existing row with the latest data instead of adding a second, inflated count.
        //  - DIFFERENT users both logging the same order is NOT blocked here — it goes in
        //    as-is, and shows up in the existing "Duplicate Order Claims" panel for
        //    management to review, username-wise, at their own pace.
        if (entryType === 'order') {
          var mine = await pool.query(
            `SELECT id FROM cs_workload_entries WHERE order_number = $1 AND entry_type = 'order' AND rep_username = $2 LIMIT 1`,
            [orderNum, repUsername]
          );
          if (mine.rows.length) {
            await pool.query(
              `UPDATE cs_workload_entries SET entry_date=$1, customer_name=$2, resolution_type=$3, booking_status=$4, fob_status=$5, file_name=$6, uploaded_at=NOW(), batch_id=$7 WHERE id=$8`,
              [entryDate, customer, resolution, bookingStatus, fobStatus, req.file.originalname, batchId, mine.rows[0].id]
            );
            updated++;
            continue;
          }
        }

        await pool.query(
          `INSERT INTO cs_workload_entries (rep_username, rep_full_name, entry_type, entry_date, order_number, customer_name, resolution_type, booking_status, fob_status, grv_reference, file_name, batch_id)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
          [repUsername, repFullName, entryType, entryDate, orderNum, customer, resolution, bookingStatus, fobStatus, grv, req.file.originalname, batchId]
        );
        inserted++;
      }

      auditLog(req.user.uid, repUsername, 'CS_WORKLOAD_UPLOAD', 'Uploaded ' + inserted + ' new + ' + updated + ' corrected ' + entryType + ' entries (' + req.file.originalname + ')', req.headers['x-forwarded-for'] || '');
      res.json({ success: true, inserted: inserted, updated: updated, batchId: batchId });
    } catch (e) {
      console.error('cs-workload upload error:', e.message);
      res.status(500).json({ error: e.message });
    }
  });

  // ── Manage Uploads: list recent individual upload batches (one row per upload
  // request) so a specific wrong file can be identified and removed on its own —
  // without a blanket Clear All that wipes everyone's data. ──
  app.get('/api/cs-workload/batches', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query(
        `SELECT batch_id, rep_username, rep_full_name, entry_type, file_name,
                MIN(uploaded_at) AS uploaded_at, COUNT(*) AS row_count
         FROM cs_workload_entries
         WHERE batch_id IS NOT NULL
         GROUP BY batch_id, rep_username, rep_full_name, entry_type, file_name
         ORDER BY MIN(uploaded_at) DESC
         LIMIT 100`
      );
      res.json({ success: true, batches: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Delete one upload batch. Superadmin/subadmin only — a rep can't delete
  // their own or anyone else's logged work, only management can. ──
  app.delete('/api/cs-workload/batches/:batchId', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var batchId = req.params.batchId;
      var check = await pool.query(
        `SELECT rep_full_name, file_name, COUNT(*) AS cnt FROM cs_workload_entries WHERE batch_id = $1 GROUP BY rep_full_name, file_name`,
        [batchId]
      );
      if (!check.rows.length) return res.status(404).json({ error: 'That upload was not found — it may have already been deleted.' });
      var r = await pool.query('DELETE FROM cs_workload_entries WHERE batch_id = $1', [batchId]);
      var info = check.rows[0];
      await auditLog(req.user.uid, req.user.username, 'CS_WORKLOAD_BATCH_DELETE',
        'Deleted upload batch (' + r.rowCount + ' rows) \u2014 ' + info.rep_full_name + ' / ' + info.file_name,
        req.headers['x-forwarded-for'] || '');
      res.json({ success: true, deleted: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
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
      // Auto vs Manual breakdown per rep (only meaningful for entry_type='order' —
      // GRV generation doesn't have an auto/manual concept)
      var ab = await pool.query(
        `SELECT rep_username, booking_status, COUNT(*) AS cnt
         FROM cs_workload_entries ${where}${where ? ' AND' : 'WHERE'} entry_type = 'order'
         GROUP BY rep_username, booking_status`, params
      );
      res.json({ success: true, rows: r.rows, autoManual: ab.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Personal view: a regular rep sees ONLY their own uploads — their daily
  // counts (for a chart), a resolution-type percentage breakdown, and their own
  // Auto vs Manual split. No visibility into anyone else's numbers, by design. ──
  async function buildUserSummary(username, repFullName, query) {
    var params = [username], dateClauses = [];
    if (query.date_from) { params.push(query.date_from); dateClauses.push('entry_date >= $' + params.length); }
    if (query.date_to) { params.push(query.date_to); dateClauses.push('entry_date <= $' + params.length); }
    var dateWhere = dateClauses.length ? (' AND ' + dateClauses.join(' AND ')) : '';

    var daily = await pool.query(
      `SELECT entry_date, entry_type, COUNT(*) AS cnt FROM cs_workload_entries
       WHERE rep_username = $1${dateWhere} GROUP BY entry_date, entry_type ORDER BY entry_date`, params
    );
    var breakdown = await pool.query(
      `SELECT COALESCE(NULLIF(resolution_type,''),'Unspecified') AS resolution_type, COUNT(*) AS cnt
       FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order'${dateWhere} GROUP BY resolution_type`, params
    );
    var totals = await pool.query(
      `SELECT entry_type, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1${dateWhere} GROUP BY entry_type`, params
    );
    var autoManual = await pool.query(
      `SELECT booking_status, COUNT(*) AS cnt FROM cs_workload_entries
       WHERE rep_username = $1 AND entry_type = 'order'${dateWhere} GROUP BY booking_status`, params
    );
    // Full 3-way FOB breakdown: AUTO BOOKED / AUTO BOOKED W/ MANUAL INTERVENTION /
    // MANUAL BOOKED — not collapsed to just Auto/Manual, so the "auto-booked but still
    // needed a person" middle category isn't lost.
    var fobBreakdown = await pool.query(
      `SELECT COALESCE(NULLIF(fob_status,''),'Unspecified') AS fob_status, COUNT(*) AS cnt
       FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order'${dateWhere} GROUP BY fob_status ORDER BY cnt DESC`, params
    );
    // Month-over-month trend (Jul, Aug, Sep...) — separate from the daily view,
    // since the daily chart alone doesn't show progress across months.
    var monthly = await pool.query(
      `SELECT to_char(entry_date, 'YYYY-MM') AS month, booking_status, COUNT(*) AS cnt
       FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order' AND entry_date IS NOT NULL${dateWhere}
       GROUP BY month, booking_status ORDER BY month`, params
    );
    return { success: true, repFullName: repFullName || username, daily: daily.rows, breakdown: breakdown.rows, totals: totals.rows, autoManual: autoManual.rows, fobBreakdown: fobBreakdown.rows, monthly: monthly.rows };
  }

  app.get('/api/cs-workload/my-summary', requireAuth, async function (req, res) {
    try {
      var result = await buildUserSummary(req.user.username, req.user.full_name, req.query);
      res.json(result);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Admin drill-down: management can view EXACTLY what a specific rep sees on
  // their own page — same shape of data, same summary — without needing that
  // rep's login. Restricted to superadmin/subadmin only. ──
  app.get('/api/cs-workload/user-summary/:username', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var target = req.params.username;
      var nameRow = await pool.query('SELECT DISTINCT rep_full_name FROM cs_workload_entries WHERE rep_username = $1 LIMIT 1', [target]);
      var fullName = nameRow.rows.length ? nameRow.rows[0].rep_full_name : target;
      var result = await buildUserSummary(target, fullName, req.query);
      res.json(result);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Personal export: a rep's own executive summary (KPIs + monthly trend +
  // resolution breakdown), colored, downloadable by the rep themselves — not
  // management-only, since each person needs their own record of their work. ──
  app.get('/api/cs-workload/my-export', requireAuth, async function (req, res) {
    try {
      var ExcelJS = require('exceljs');
      var NAVY = 'FF1B2338', WHITE = 'FFFFFFFF', GOLD = 'FFFCF3CF', GREEN = 'FFE1F5E9';
      var username = req.user.username;
      var fullName = req.user.full_name || username;

      var totals = await pool.query(`SELECT entry_type, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1 GROUP BY entry_type`, [username]);
      var autoManual = await pool.query(`SELECT booking_status, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order' GROUP BY booking_status`, [username]);
      var monthly = await pool.query(`SELECT to_char(entry_date, 'YYYY-MM') AS month, booking_status, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order' AND entry_date IS NOT NULL GROUP BY month, booking_status ORDER BY month`, [username]);
      var breakdown = await pool.query(`SELECT COALESCE(NULLIF(resolution_type,''),'Unspecified') AS resolution_type, COUNT(*) AS cnt FROM cs_workload_entries WHERE rep_username = $1 AND entry_type = 'order' GROUP BY resolution_type`, [username]);

      var totalOrders = 0, totalGrv = 0;
      totals.rows.forEach(function (t) { if (t.entry_type === 'order') totalOrders = +t.cnt; if (t.entry_type === 'grv') totalGrv = +t.cnt; });
      var autoCount = 0, manualCount = 0;
      autoManual.rows.forEach(function (a) { if (a.booking_status === 'Auto') autoCount = +a.cnt; else manualCount = +a.cnt; });

      var wb = new ExcelJS.Workbook();
      var es = wb.addWorksheet('My Executive Summary');
      es.columns = [{ width: 26 }, { width: 16 }, { width: 16 }];
      var t = es.addRow([fullName.toUpperCase() + ' \u2014 EXECUTIVE SUMMARY']);
      es.mergeCells('A' + t.number + ':C' + t.number);
      t.font = { bold: true, size: 13, color: { argb: NAVY } };
      es.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      es.addRow([]);
      var hdr = es.addRow(['Metric', 'Value', '']);
      hdr.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      [['Total Orders Processed', totalOrders], ['Auto-Booked', autoCount], ['Manual', manualCount],
       ['Manual %', totalOrders ? Math.round(manualCount / totalOrders * 1000) / 10 + '%' : '0%'], ['GRV Generated', totalGrv]]
        .forEach(function (row) {
          var r = es.addRow(row);
          r.getCell(1).font = { bold: true };
          r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
        });
      es.addRow([]);

      var monthMap = {};
      monthly.rows.forEach(function (row) {
        if (!monthMap[row.month]) monthMap[row.month] = { auto: 0, manual: 0 };
        if (row.booking_status === 'Auto') monthMap[row.month].auto = +row.cnt; else monthMap[row.month].manual = +row.cnt;
      });
      var mHdr = es.addRow(['Month', 'Auto', 'Manual', 'Total']);
      mHdr.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      Object.keys(monthMap).sort().forEach(function (m) {
        var v = monthMap[m];
        var r = es.addRow([m, v.auto, v.manual, v.auto + v.manual]);
        r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } }; });
      });
      es.addRow([]);

      var bHdr = es.addRow(['Resolution Type', 'Count']);
      bHdr.eachCell(function (c) { c.font = { bold: true, color: { argb: WHITE } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; });
      breakdown.rows.forEach(function (row) { es.addRow([row.resolution_type, +row.cnt]); });

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="My_Executive_Summary_' + username + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('cs-workload my-export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
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
        var nonBlankCount = grid[i].filter(function (c) { return String(c).trim() !== ''; }).length;
        // A real header row has several populated columns. A single descriptive
        // sentence cell (e.g. a section title that happens to mention "person" and
        // "manual" in the same breath) only has one — skip those.
        if (nonBlankCount < 3) continue;
        var pIdx = nr.findIndex(function (c) { return c.indexOf('person') !== -1 || c.indexOf('rep') !== -1 || c.indexOf('cst') !== -1; });
        var mIdx = nr.findIndex(function (c) { return c.indexOf('manual') !== -1; });
        if (pIdx !== -1 && mIdx !== -1 && pIdx !== mIdx) {
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
      if (headerRowIdx === -1) return res.status(400).json({ error: 'Could not find a header row with Person and Manual columns in that file. Sheet read: "' + wb.SheetNames[0] + '".' });

      // Hard stop instead of silently inserting zeros — if Total/Auto/Manual weren't
      // actually found as real columns, this is very likely the wrong file/sheet
      // (e.g. a customer detail list, not the Person/Total/Auto/Manual summary table).
      var missing = [];
      if (C.total === -1) missing.push('Total');
      if (C.auto === -1) missing.push('Auto');
      if (C.manual === -1 || C.manual === undefined) missing.push('Manual');
      if (missing.length) {
        return res.status(400).json({
          error: 'Found a "Person" column but no ' + missing.join('/') + ' column(s) in sheet "' + wb.SheetNames[0] + '". ' +
            'Header row detected: [' + grid[headerRowIdx].filter(function (c) { return String(c).trim() !== ''; }).join(', ') + ']. ' +
            'This looks like the wrong sheet/file \u2014 upload the summary sheet with Person, Total, Auto, and Manual columns.'
        });
      }

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
  // Personal version: any rep sees duplicates involving THEIR OWN uploads — persistent
  // on their own page, so "Ameer and you both added this order" is visible whenever they
  // load the page, not just as a one-time popup right after uploading.
  app.get('/api/cs-workload/my-duplicates', requireAuth, async function (req, res) {
    try {
      var username = req.user.username;
      var mine = await pool.query(
        `SELECT DISTINCT order_number FROM cs_workload_entries
         WHERE rep_username = $1 AND entry_type = 'order' AND order_number IS NOT NULL AND order_number != ''`,
        [username]
      );
      if (!mine.rows.length) return res.json({ success: true, duplicates: [] });
      var orderNums = mine.rows.map(function (r) { return r.order_number; });
      var shared = await pool.query(
        `SELECT order_number FROM cs_workload_entries
         WHERE entry_type = 'order' AND order_number = ANY($1)
         GROUP BY order_number HAVING COUNT(DISTINCT rep_username) > 1`, [orderNums]
      );
      if (!shared.rows.length) return res.json({ success: true, duplicates: [] });
      var sharedNums = shared.rows.map(function (r) { return r.order_number; });
      var detail = await pool.query(
        `SELECT order_number, rep_username, rep_full_name, entry_date, customer_name, resolution_type, uploaded_at
         FROM cs_workload_entries WHERE entry_type = 'order' AND order_number = ANY($1)
         ORDER BY order_number, uploaded_at`, [sharedNums]
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
      var autoManualR = await pool.query(
        `SELECT rep_username, booking_status, COUNT(*) AS cnt FROM cs_workload_entries
         ${where}${where ? ' AND' : 'WHERE'} entry_type = 'order' GROUP BY rep_username, booking_status`, params
      );
      // Day-wise tracking: one row per Rep + Booked Date, with Auto/Manual counts and %
      // — full day-by-day coverage, not just a single aggregate total.
      var dayWiseR = await pool.query(
        `SELECT rep_username, rep_full_name, entry_date, booking_status, COUNT(*) AS cnt
         FROM cs_workload_entries ${where}${where ? ' AND' : 'WHERE'} entry_type = 'order' AND entry_date IS NOT NULL
         GROUP BY rep_username, rep_full_name, entry_date, booking_status ORDER BY entry_date, rep_full_name`, params
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

      // ── Sheet 1: Scorecard summary, now with Auto/Manual counts and % ──
      var sc = wb.addWorksheet('Scorecard');
      sc.columns = [{ width: 22 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 18 }];
      var t = sc.addRow(['CS WORKLOAD SCORECARD']);
      sc.mergeCells('A' + t.number + ':F' + t.number);
      t.font = { bold: true, size: 13, color: { argb: NAVY } };
      sc.addRow(['Period', (req.query.date_from || 'All time') + ' to ' + (req.query.date_to || 'Latest')]);
      sc.addRow([]);
      headerRow(sc, ['Rep', 'Orders', 'Auto', 'Manual', 'Manual %', 'Distinct Customers']);
      var orderRows = scoreR.rows.filter(function (r) { return r.entry_type === 'order'; });
      var amByUser = {};
      autoManualR.rows.forEach(function (r) {
        if (!amByUser[r.rep_username]) amByUser[r.rep_username] = { auto: 0, manual: 0 };
        if (r.booking_status === 'Auto') amByUser[r.rep_username].auto = +r.cnt; else amByUser[r.rep_username].manual = +r.cnt;
      });
      orderRows.forEach(function (row) {
        var am = amByUser[row.rep_username] || { auto: 0, manual: 0 };
        var total = +row.cnt;
        var manualPct = total ? Math.round(am.manual / total * 1000) / 10 : 0;
        sc.addRow([row.rep_full_name || row.rep_username, total, am.auto, am.manual, manualPct + '%', +row.distinct_customers]);
      });

      // ── Sheet 2: Day-Wise Tracking — full daily coverage per rep, Auto/Manual/% ──
      var dw = wb.addWorksheet('Day-Wise Tracking');
      dw.columns = [{ width: 14 }, { width: 22 }, { width: 10 }, { width: 10 }, { width: 12 }, { width: 10 }];
      var t3 = dw.addRow(['DAY-WISE TRACKING \u2014 100% DAILY COVERAGE PER REP']);
      dw.mergeCells('A' + t3.number + ':F' + t3.number);
      t3.font = { bold: true, size: 12, color: { argb: NAVY } };
      dw.addRow([]);
      headerRow(dw, ['Booked Date', 'Rep', 'Auto', 'Manual', 'Manual %', 'Total']);
      var byDayRep = {};
      dayWiseR.rows.forEach(function (row) {
        var key = row.entry_date + '|' + row.rep_username;
        if (!byDayRep[key]) byDayRep[key] = { date: row.entry_date, rep: row.rep_full_name || row.rep_username, auto: 0, manual: 0 };
        if (row.booking_status === 'Auto') byDayRep[key].auto = +row.cnt; else byDayRep[key].manual = +row.cnt;
      });
      var dayKeys = Object.keys(byDayRep).sort(function (a, b) { return new Date(byDayRep[a].date) - new Date(byDayRep[b].date); });
      dayKeys.forEach(function (k) {
        var v = byDayRep[k];
        var total = v.auto + v.manual;
        var manualPct = total ? Math.round(v.manual / total * 1000) / 10 : 0;
        var row = dw.addRow([new Date(v.date).toLocaleDateString('en-AE'), v.rep, v.auto, v.manual, manualPct + '%', total]);
        row.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
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

  // ── AI Insights: plain-language commentary on the current workload picture.
  // Superadmin/subadmin only — this is a management tool, not something reps see.
  app.get('/api/cs-workload/ai-comments', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var baseline = await pool.query('SELECT person_name, account_label, total_orders, auto_orders, manual_orders FROM cs_workload_baseline ORDER BY manual_orders DESC');
      if (!baseline.rows.length) return res.json({ success: true, comment: 'No baseline data uploaded yet \u2014 upload the 3-Month Historical Scorecard first, then ask again.' });

      var teamManualTotal = baseline.rows.reduce(function (s, r) { return s + (+r.manual_orders); }, 0);
      var dataText = baseline.rows.map(function (r) {
        var total = +r.total_orders, manual = +r.manual_orders;
        var manualPct = total ? Math.round(manual / total * 1000) / 10 : 0;
        var share = teamManualTotal ? Math.round(manual / teamManualTotal * 1000) / 10 : 0;
        return r.person_name + ': total=' + total + ', manual=' + manual + ' (' + manualPct + '% of their own orders), ' + share + '% of team-wide manual workload';
      }).join('\n');

      var prompt = 'You are advising a UAE food distribution company\'s operations manager on customer-service workload balance. ' +
        'Below is real per-person data: total orders, how many needed manual handling (vs auto-booked), and each person\'s share of the whole team\'s manual workload.\n\n' +
        dataText + '\n\n' +
        'Give a short, direct set of observations (max 150 words, plain text, no markdown headers) covering: (1) who is carrying the most manual load and by how much relative to others, ' +
        '(2) anyone whose manual % of their own orders is unusually high (meaning their work is inherently harder, not just higher volume), ' +
        '(3) anyone at or near zero who may need a real assignment. Be specific with the numbers given, not generic.';

      var msg = await anthropic.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 600,
        messages: [{ role: 'user', content: prompt }]
      });
      var comment = (msg.content && msg.content[0] && msg.content[0].text) ? msg.content[0].text.trim() : 'No response generated.';
      res.json({ success: true, comment: comment });
    } catch (e) {
      console.error('cs-workload ai-comments error:', e.message);
      res.status(500).json({ error: 'Could not generate insights right now. Please try again in a moment.' });
    }
  });

  app.delete('/api/cs-workload/clear', requireAuth, requireRole('superadmin'), async function (req, res) {
    try {
      var r1 = await pool.query('DELETE FROM cs_workload_entries');
      var r2 = await pool.query('DELETE FROM cs_workload_baseline');
      var total = r1.rowCount + r2.rowCount;
      await auditLog(req.user.uid, req.user.username, 'CS_WORKLOAD_CLEAR', 'Cleared all CS workload data \u2014 ' + r1.rowCount + ' entries + ' + r2.rowCount + ' baseline rows', req.headers['x-forwarded-for'] || '');
      res.json({ success: true, deleted: total, entriesDeleted: r1.rowCount, baselineDeleted: r2.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
};
