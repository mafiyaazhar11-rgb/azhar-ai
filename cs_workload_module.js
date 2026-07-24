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
        order: fc(keys, ['ordernumber', 'order']),
        customer: fc(keys, ['customer']),
        resolution: fc(keys, ['resolution', 'reason', 'type']),
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
  app.get('/api/cs-workload/scorecard', requireAuth, async function (req, res) {
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

  // ── Duplicate order claims: same order number logged by more than one rep ──
  app.get('/api/cs-workload/duplicates', requireAuth, async function (req, res) {
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
  app.get('/api/cs-workload/export', requireAuth, async function (req, res) {
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
