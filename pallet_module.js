// ============================================================
// AUJAN PALLET COLLECTION & RECOVERY TRACKING MODULE — Phase 1
// Plugs into existing AZHAR-AI server.js. Self-contained: own tables,
// shares the existing pg pool + auth.
// Mount with: require('./pallet_module')(app, pool, requireAuth, requireRole, upload, auditLog);
//
// PHASE 1 SCOPE (per requirement doc, section 18):
//   - Excel/CSV upload with validation + duplicate detection
//   - Pallet Cost Master (admin-editable, default AED 48/pallet)
//   - Executive KPI dashboard
//   - Customer-wise, Door-wise, Salesman-wise reports + Ageing analysis
//   - Status field (tracked, but full workflow-history UI is Phase 2)
//   - Excel download for every report
// Phase 2 (action ownership/escalation, transport workflow, charge workflow,
// alerts, executive summary doc, user access controls) and Phase 3 (Azhar AI
// analysis, Qlik, email automation, mobile, trend history) are intentionally
// NOT built here — tables are shaped so Phase 2/3 can extend without a
// schema rewrite (status_history and department/priority columns already
// exist), but the workflow logic itself is deferred as requested.
//
// Customer/Site linkage: this module does NOT create a second customer
// master. It soft-links customer_code -> horeca_customer_sites.account_number
// and customer_site_id -> horeca_customer_sites.site_use_id (the existing
// master in this app) purely for lookups/autofill — pallet uploads are never
// blocked by a missing master record, since Aujan customers may not all
// exist in the HoReCa master yet.
// ============================================================

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog) {

  const GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', REDFLAG = 'FFFDE0DE';

  const STATUS_LIST = ['Delivered', 'Pending customer confirmation', 'Collection requested',
    'Collection planned', 'Collection attempted', 'Partially collected', 'Fully collected',
    'Customer refused', 'Pallet not available', 'Under commercial review',
    'Charge approved', 'Charged to customer', 'Closed'];

  const DEPARTMENTS = ['Customer Service', 'Warehouse', 'Transport', 'Sales', 'Commercial', 'Finance'];
  const PRIORITIES = ['Normal', 'High', 'Critical'];

  // Excel/CSV upload template header -> internal field name.
  // NOTE: "Outstanding Pallets" is intentionally excluded from the upload
  // template — per section 3 of the requirement it is a calculated field,
  // never a manually-entered one, so we don't give users a place to
  // contradict the calculation.
  const HEADER_MAP = [
    ['Delivery Date', 'delivery_date', 'date'],
    ['Invoice Number', 'invoice_number', 'text'],
    ['Sales Order Number', 'sales_order_number', 'text'],
    ['Customer Code', 'customer_code', 'text'],
    ['Customer Name', 'customer_name', 'text'],
    ['Customer Site ID', 'customer_site_id', 'text'],
    ['Ship-To Location', 'ship_to_location', 'text'],
    ['Customer Door', 'customer_door', 'text'],
    ['Emirate / Region', 'region', 'text'],
    ['Salesman Name', 'salesman_name', 'text'],
    ['Sales Supervisor', 'sales_supervisor', 'text'],
    ['Brand', 'brand', 'text'],
    ['Pallet Type', 'pallet_type', 'text'],
    ['Pallets Delivered', 'pallets_delivered', 'number'],
    ['Pallets Returned', 'pallets_returned', 'number'],
    ['Return Date', 'return_date', 'date'],
    ['Driver Name', 'driver_name', 'text'],
    ['Vehicle Number', 'vehicle_number', 'text'],
    ['Transport Type (In-house/3PL)', 'transport_type', 'text'],
    ['Collection Request Date', 'collection_request_date', 'date'],
    ['Collection Attempt Date', 'collection_attempt_date', 'date'],
    ['Collection Status', 'collection_status', 'text'],
    ['Customer Confirmation', 'customer_confirmation', 'text'],
    ['Reason Not Collected', 'reason_not_collected', 'text'],
    ['Next Action Date', 'next_action_date', 'date'],
    ['Action Owner', 'action_owner', 'text'],
    ['Charge Applicable (Yes/No)', 'charge_applicable', 'text'],
    ['Charge Amount', 'charge_amount', 'number'],
    ['Charge Status', 'charge_status', 'text'],
    ['Remarks', 'remarks', 'text']
  ];
  const MANDATORY_FIELDS = ['delivery_date', 'invoice_number', 'customer_code', 'customer_site_id', 'customer_name', 'customer_door', 'pallets_delivered'];

  // Cost lookup used everywhere a value is calculated: exact pallet_type match,
  // falling back to 'Standard', falling back to the hardcoded AED 48 default.
  const UNIT_COST_SQL = `
    COALESCE(
      (SELECT cost_per_pallet FROM pallet_cost_master m WHERE m.pallet_type = t.pallet_type AND m.active = true AND m.effective_from <= CURRENT_DATE ORDER BY m.effective_from DESC, m.id DESC LIMIT 1),
      (SELECT cost_per_pallet FROM pallet_cost_master m WHERE m.pallet_type = 'Standard' AND m.active = true AND m.effective_from <= CURRENT_DATE ORDER BY m.effective_from DESC, m.id DESC LIMIT 1),
      48
    )`;

  // Reusable computed-columns block. `t` must be the alias of pallet_transactions.
  const CALC_COLUMNS = `
      GREATEST(t.pallets_delivered - t.pallets_returned, 0) AS outstanding_pallets,
      CASE WHEN t.pallets_delivered > 0 THEN ROUND((t.pallets_returned / t.pallets_delivered * 100)::numeric, 1) ELSE 0 END AS recovery_pct,
      CASE
        WHEN t.pallets_returned >= t.pallets_delivered AND t.return_date IS NOT NULL THEN (t.return_date - t.delivery_date)
        ELSE (CURRENT_DATE - t.delivery_date)
      END AS outstanding_days,
      (${UNIT_COST_SQL}) AS unit_cost,
      GREATEST(t.pallets_delivered - t.pallets_returned, 0) * (${UNIT_COST_SQL}) AS value_at_risk`;

  function ageingBucketSQL(daysExpr) {
    return `CASE
      WHEN ${daysExpr} <= 7 THEN '0-7 days'
      WHEN ${daysExpr} <= 15 THEN '8-15 days'
      WHEN ${daysExpr} <= 30 THEN '16-30 days'
      WHEN ${daysExpr} <= 60 THEN '31-60 days'
      ELSE 'Above 60 days'
    END`;
  }

  // ── Init tables ──
  async function initPalletDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS pallet_cost_master (
        id SERIAL PRIMARY KEY,
        pallet_type TEXT NOT NULL,
        cost_per_pallet NUMERIC NOT NULL,
        effective_from DATE NOT NULL DEFAULT CURRENT_DATE,
        active BOOLEAN NOT NULL DEFAULT true,
        updated_by TEXT,
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_cost_type ON pallet_cost_master(pallet_type, active, effective_from DESC)`);
      var seed = await pool.query(`SELECT id FROM pallet_cost_master WHERE pallet_type = 'Standard'`);
      if (!seed.rows.length) {
        await pool.query(
          `INSERT INTO pallet_cost_master (pallet_type, cost_per_pallet, effective_from, active, updated_by) VALUES ($1,$2,CURRENT_DATE,true,$3)`,
          ['Standard', 48, 'system']
        );
      }

      await pool.query(`CREATE TABLE IF NOT EXISTS pallet_transactions (
        id SERIAL PRIMARY KEY,
        delivery_date DATE NOT NULL,
        invoice_number TEXT NOT NULL,
        sales_order_number TEXT,
        customer_code TEXT NOT NULL,
        customer_name TEXT NOT NULL,
        customer_site_id TEXT NOT NULL,
        ship_to_location TEXT,
        customer_door TEXT NOT NULL,
        region TEXT,
        salesman_name TEXT,
        sales_supervisor TEXT,
        brand TEXT DEFAULT 'Aujan',
        pallet_type TEXT NOT NULL DEFAULT 'Standard',
        pallets_delivered NUMERIC NOT NULL DEFAULT 0,
        pallets_returned NUMERIC NOT NULL DEFAULT 0,
        return_date DATE,
        driver_name TEXT,
        vehicle_number TEXT,
        transport_type TEXT,
        collection_request_date DATE,
        collection_attempt_date DATE,
        collection_status TEXT NOT NULL DEFAULT 'Delivered',
        customer_confirmation TEXT,
        reason_not_collected TEXT,
        next_action_date DATE,
        action_owner TEXT,
        action_department TEXT,
        priority TEXT NOT NULL DEFAULT 'Normal',
        charge_applicable TEXT DEFAULT 'No',
        charge_amount NUMERIC,
        charge_status TEXT,
        remarks TEXT,
        flags JSONB NOT NULL DEFAULT '[]'::jsonb,
        upload_batch_id TEXT,
        created_by INT REFERENCES users(id),
        created_by_name TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_by INT REFERENCES users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(invoice_number, customer_site_id, pallet_type)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_customer ON pallet_transactions(customer_code)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_delivery ON pallet_transactions(delivery_date)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_status ON pallet_transactions(collection_status)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_salesman ON pallet_transactions(salesman_name)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_door ON pallet_transactions(customer_door)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_txn_batch ON pallet_transactions(upload_batch_id)`);

      // Phase 2 will build full workflow UI on top of this; the table exists
      // now so Phase 1 status changes (if any happen via re-upload) aren't lost.
      await pool.query(`CREATE TABLE IF NOT EXISTS pallet_status_history (
        id SERIAL PRIMARY KEY,
        transaction_id INT REFERENCES pallet_transactions(id) ON DELETE CASCADE,
        previous_status TEXT,
        new_status TEXT,
        updated_by TEXT,
        comment TEXT,
        next_action_date DATE,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_status_hist_txn ON pallet_status_history(transaction_id, created_at DESC)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS pallet_uploads_log (
        id SERIAL PRIMARY KEY,
        batch_id TEXT NOT NULL,
        file_name TEXT,
        uploaded_by TEXT,
        uploaded_by_id INT,
        total_rows INT DEFAULT 0,
        inserted_rows INT DEFAULT 0,
        updated_rows INT DEFAULT 0,
        rejected_rows INT DEFAULT 0,
        duplicate_rows INT DEFAULT 0,
        validation_errors JSONB DEFAULT '[]'::jsonb,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_pallet_uploads_created ON pallet_uploads_log(created_at DESC)`);

      // ── Aujan Supply Tracking (warehouse-level, separate from customer flow) ──
      // Two upload types share one table, split by entry_type: 'inbound' = pallets
      // received FROM Aujan into the warehouse; 'return' = pallets sent BACK to
      // Aujan. Shipment-level, not per-customer — e.g. "5 trucks, 200 pallets
      // received today" or "100 pallets returned to Aujan today".
      await pool.query(`CREATE TABLE IF NOT EXISTS aujan_supply_log (
        id SERIAL PRIMARY KEY,
        entry_type TEXT NOT NULL CHECK (entry_type IN ('inbound','return')),
        entry_date DATE NOT NULL,
        reference TEXT,
        truck_count NUMERIC,
        pallet_type TEXT NOT NULL DEFAULT 'Standard',
        pallet_qty NUMERIC NOT NULL DEFAULT 0,
        warehouse TEXT,
        remarks TEXT,
        upload_batch_id TEXT,
        created_by INT REFERENCES users(id),
        created_by_name TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_supply_log_type_date ON aujan_supply_log(entry_type, entry_date)`);

      console.log('Pallet module DB ready');
    } catch (e) {
      console.error('Pallet DB init error:', e.message);
    }
  }
  initPalletDB();

  // ── Helpers ──

  function toDateStr(val) {
    if (val === undefined || val === null || val === '') return null;
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return null;
      return val.toISOString().slice(0, 10);
    }
    if (typeof val === 'number') {
      // Excel serial date fallback (in case cellDates:true didn't catch it)
      var d = new Date(Math.round((val - 25569) * 86400 * 1000));
      return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
    }
    var parsed = new Date(val);
    return isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
  }

  function toNum(val) {
    if (val === undefined || val === null || val === '') return 0;
    var n = Number(String(val).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  function buildFilters(q) {
    var clauses = [];
    var params = [];
    var p = function (v) { params.push(v); return '$' + params.length; };

    if (q.date_from) clauses.push('t.delivery_date >= ' + p(q.date_from));
    if (q.date_to) clauses.push('t.delivery_date <= ' + p(q.date_to));
    if (q.customer_code) clauses.push('t.customer_code = ' + p(q.customer_code));
    if (q.customer_site_id) clauses.push('t.customer_site_id = ' + p(q.customer_site_id));
    if (q.customer_door) clauses.push('t.customer_door = ' + p(q.customer_door));
    if (q.region) clauses.push('t.region = ' + p(q.region));
    if (q.salesman_name) clauses.push('t.salesman_name = ' + p(q.salesman_name));
    if (q.sales_supervisor) clauses.push('t.sales_supervisor = ' + p(q.sales_supervisor));
    if (q.brand) clauses.push('t.brand = ' + p(q.brand));
    if (q.driver_name) clauses.push('t.driver_name = ' + p(q.driver_name));
    if (q.transport_type) clauses.push('t.transport_type = ' + p(q.transport_type));
    if (q.collection_status) clauses.push('t.collection_status = ' + p(q.collection_status));
    if (q.charge_status) clauses.push('t.charge_status = ' + p(q.charge_status));
    if (q.action_owner) clauses.push('t.action_owner = ' + p(q.action_owner));
    if (q.ageing_bucket) {
      clauses.push(ageingBucketSQL('(CURRENT_DATE - t.delivery_date)') + ' = ' + p(q.ageing_bucket));
    }
    if (q.outstanding_only === 'true') clauses.push('(t.pallets_delivered - t.pallets_returned) > 0');

    return { where: clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '', params: params };
  }

  function filterSummaryStr(q) {
    var parts = [];
    if (q.date_from || q.date_to) parts.push('Date: ' + (q.date_from || 'earliest') + ' to ' + (q.date_to || 'today'));
    ['customer_code', 'customer_door', 'region', 'salesman_name', 'brand', 'transport_type', 'collection_status', 'ageing_bucket'].forEach(function (k) {
      if (q[k]) parts.push(k.replace(/_/g, ' ') + ': ' + q[k]);
    });
    return parts.length ? parts.join(' | ') : 'No filters (all data)';
  }

  // ── Excel styling helpers (matches existing app convention) ──
  function styleHeaderRow(row) {
    row.eachCell(function (cell) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
      cell.font = { bold: true, color: { argb: GOLD }, size: 11 };
      cell.alignment = { vertical: 'middle' };
    });
  }
  function styleTotalRow(row) {
    row.eachCell(function (cell) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHTGOLD } };
      cell.font = { bold: true, color: { argb: DARKBG } };
    });
  }
  function styleTitleRow(ws, text, span) {
    var r = ws.addRow([text]);
    r.font = { bold: true, size: 15, color: { argb: GOLD } };
    ws.mergeCells('A' + r.number + ':' + span + r.number);
    ws.getRow(r.number).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
  }

  // ── ADMIN: Pallet Cost Master ──

  app.get('/api/pallets/cost-master', requireAuth, async function (req, res) {
    try {
      var r = await pool.query('SELECT * FROM pallet_cost_master ORDER BY pallet_type, effective_from DESC');
      res.json({ rates: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/pallets/cost-master', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var b = req.body || {};
      if (!b.pallet_type || b.cost_per_pallet === undefined) return res.status(400).json({ error: 'pallet_type and cost_per_pallet are required' });
      var r = await pool.query(
        `INSERT INTO pallet_cost_master (pallet_type, cost_per_pallet, effective_from, active, updated_by)
         VALUES ($1,$2,COALESCE($3,CURRENT_DATE),true,$4) RETURNING *`,
        [b.pallet_type, toNum(b.cost_per_pallet), b.effective_from || null, req.user.username]
      );
      await auditLog(req.user.uid, req.user.username, 'PALLET_COST_UPDATE', b.pallet_type + ' -> AED ' + b.cost_per_pallet, req.headers['x-forwarded-for'] || '');
      res.json({ rate: r.rows[0] });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.put('/api/pallets/cost-master/:id/deactivate', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('UPDATE pallet_cost_master SET active=false, updated_by=$1, updated_at=NOW() WHERE id=$2 RETURNING *', [req.user.username, req.params.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Rate not found' });
      res.json({ rate: r.rows[0] });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── UPLOAD TEMPLATE ──

  app.get('/api/pallets/upload-template', requireAuth, async function (req, res) {
    try {
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI';
      wb.created = new Date();
      var ws = wb.addWorksheet('Pallet Data');
      ws.columns = HEADER_MAP.map(function () { return { width: 20 }; });
      var headerRow = ws.addRow(HEADER_MAP.map(function (h) { return h[0]; }));
      styleHeaderRow(headerRow);
      var note = ws.addRow(['Mandatory: Delivery Date, Invoice Number, Customer Code, Customer Site ID, Customer Name, Customer Door, Pallets Delivered. Dates as DD-MM-YYYY or YYYY-MM-DD. Outstanding pallets, Recovery %, Ageing and Value at Risk are calculated automatically — do not enter them.']);
      ws.mergeCells('A' + (headerRow.number) + ':' + 'A' + headerRow.number); // no-op, keeps layout stable
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_Pallet_Upload_Template.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── UPLOAD DATA (Excel or CSV) ──

  app.post('/api/pallets/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try {
        wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      } catch (parseErr) {
        return res.status(400).json({ error: 'Could not read that file. Upload the .xlsx template or a CSV export of it.' });
      }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      var rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rawRows.length) return res.status(400).json({ error: 'The file has no data rows.' });

      var batchId = 'PLT-' + Date.now();
      var inserted = 0, updated = 0, rejected = 0, duplicates = 0;
      var errors = [];
      var seenInBatch = {}; // key -> row# for in-file duplicate detection

      for (var i = 0; i < rawRows.length; i++) {
        var raw = rawRows[i];
        var rowNum = i + 2; // header is row 1
        var rec = {};
        HEADER_MAP.forEach(function (h) {
          var display = h[0], field = h[1], type = h[2];
          var val = raw[display];
          if (val === undefined) {
            // tolerate minor header variance (case/spacing)
            var key = Object.keys(raw).find(function (k) { return k.trim().toLowerCase() === display.trim().toLowerCase(); });
            if (key) val = raw[key];
          }
          if (type === 'date') rec[field] = toDateStr(val);
          else if (type === 'number') rec[field] = toNum(val);
          else rec[field] = (val === undefined || val === null) ? null : String(val).trim();
        });

        // Mandatory field validation
        var missing = MANDATORY_FIELDS.filter(function (f) {
          return rec[f] === null || rec[f] === '' || (f === 'pallets_delivered' && (rec[f] === undefined));
        });
        if (missing.length) {
          rejected++;
          errors.push({ row: rowNum, invoice: rec.invoice_number || '', reason: 'Missing mandatory field(s): ' + missing.join(', ') });
          continue;
        }

        var dupKey = rec.invoice_number + '|' + rec.customer_site_id + '|' + (rec.pallet_type || 'Standard');
        if (seenInBatch[dupKey]) {
          duplicates++;
          errors.push({ row: rowNum, invoice: rec.invoice_number, reason: 'Duplicate within this file (same invoice + site + pallet type as row ' + seenInBatch[dupKey] + ')' });
          continue;
        }
        seenInBatch[dupKey] = rowNum;

        // Data-quality flags (non-blocking, shown as alerts on the dashboard)
        var flags = [];
        if (!rec.customer_site_id) flags.push('missing_site_id');
        if (!rec.salesman_name) flags.push('missing_salesman');
        if (!rec.customer_door) flags.push('missing_door');
        if (toNum(rec.pallets_returned) > toNum(rec.pallets_delivered)) flags.push('returned_exceeds_delivered');

        rec.pallet_type = rec.pallet_type || 'Standard';
        rec.collection_status = STATUS_LIST.includes(rec.collection_status) ? rec.collection_status : (rec.collection_status ? rec.collection_status : 'Delivered');
        rec.charge_applicable = rec.charge_applicable && rec.charge_applicable.toLowerCase() === 'yes' ? 'Yes' : 'No';

        try {
          var existing = await pool.query(
            'SELECT id, collection_status FROM pallet_transactions WHERE invoice_number=$1 AND customer_site_id=$2 AND pallet_type=$3',
            [rec.invoice_number, rec.customer_site_id, rec.pallet_type]
          );

          if (existing.rows.length) {
            // Already in the system from a prior upload — update rather than
            // silently skip, since re-uploads are how W/H reports returns over time.
            await pool.query(
              `UPDATE pallet_transactions SET
                sales_order_number=$1, customer_name=$2, ship_to_location=$3, customer_door=$4, region=$5,
                salesman_name=$6, sales_supervisor=$7, brand=$8, pallets_delivered=$9, pallets_returned=$10,
                return_date=$11, driver_name=$12, vehicle_number=$13, transport_type=$14,
                collection_request_date=$15, collection_attempt_date=$16, collection_status=$17,
                customer_confirmation=$18, reason_not_collected=$19, next_action_date=$20, action_owner=$21,
                charge_applicable=$22, charge_amount=$23, charge_status=$24, remarks=$25, flags=$26,
                upload_batch_id=$27, updated_by=$28, updated_at=NOW()
               WHERE id=$29`,
              [rec.sales_order_number, rec.customer_name, rec.ship_to_location, rec.customer_door, rec.region,
              rec.salesman_name, rec.sales_supervisor, rec.brand, rec.pallets_delivered, rec.pallets_returned,
              rec.return_date, rec.driver_name, rec.vehicle_number, rec.transport_type,
              rec.collection_request_date, rec.collection_attempt_date, rec.collection_status,
              rec.customer_confirmation, rec.reason_not_collected, rec.next_action_date, rec.action_owner,
              rec.charge_applicable, rec.charge_amount || null, rec.charge_status, rec.remarks, JSON.stringify(flags),
              batchId, req.user.uid, existing.rows[0].id]
            );
            if (existing.rows[0].collection_status !== rec.collection_status) {
              await pool.query(
                `INSERT INTO pallet_status_history (transaction_id, previous_status, new_status, updated_by, comment, next_action_date)
                 VALUES ($1,$2,$3,$4,$5,$6)`,
                [existing.rows[0].id, existing.rows[0].collection_status, rec.collection_status, req.user.username, 'Updated via re-upload (batch ' + batchId + ')', rec.next_action_date]
              );
            }
            updated++;
          } else {
            var ins = await pool.query(
              `INSERT INTO pallet_transactions (
                delivery_date, invoice_number, sales_order_number, customer_code, customer_name, customer_site_id,
                ship_to_location, customer_door, region, salesman_name, sales_supervisor, brand, pallet_type,
                pallets_delivered, pallets_returned, return_date, driver_name, vehicle_number, transport_type,
                collection_request_date, collection_attempt_date, collection_status, customer_confirmation,
                reason_not_collected, next_action_date, action_owner, charge_applicable, charge_amount,
                charge_status, remarks, flags, upload_batch_id, created_by, created_by_name
              ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33,$34)
              RETURNING id`,
              [rec.delivery_date, rec.invoice_number, rec.sales_order_number, rec.customer_code, rec.customer_name, rec.customer_site_id,
              rec.ship_to_location, rec.customer_door, rec.region, rec.salesman_name, rec.sales_supervisor, rec.brand, rec.pallet_type,
              rec.pallets_delivered, rec.pallets_returned, rec.return_date, rec.driver_name, rec.vehicle_number, rec.transport_type,
              rec.collection_request_date, rec.collection_attempt_date, rec.collection_status, rec.customer_confirmation,
              rec.reason_not_collected, rec.next_action_date, rec.action_owner, rec.charge_applicable, rec.charge_amount || null,
              rec.charge_status, rec.remarks, JSON.stringify(flags), batchId, req.user.uid, req.user.full_name || req.user.username]
            );
            await pool.query(
              `INSERT INTO pallet_status_history (transaction_id, previous_status, new_status, updated_by, comment, next_action_date)
               VALUES ($1,NULL,$2,$3,$4,$5)`,
              [ins.rows[0].id, rec.collection_status, req.user.username, 'Created via upload (batch ' + batchId + ')', rec.next_action_date]
            );
            inserted++;
          }
        } catch (rowErr) {
          rejected++;
          errors.push({ row: rowNum, invoice: rec.invoice_number, reason: rowErr.message });
        }
      }

      await pool.query(
        `INSERT INTO pallet_uploads_log (batch_id, file_name, uploaded_by, uploaded_by_id, total_rows, inserted_rows, updated_rows, rejected_rows, duplicate_rows, validation_errors)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`,
        [batchId, req.file.originalname, req.user.username, req.user.uid, rawRows.length, inserted, updated, rejected, duplicates, JSON.stringify(errors)]
      );
      await auditLog(req.user.uid, req.user.username, 'PALLET_UPLOAD', req.file.originalname + ' — inserted ' + inserted + ', updated ' + updated + ', rejected ' + rejected, req.headers['x-forwarded-for'] || '');

      res.json({
        batch_id: batchId,
        total_rows: rawRows.length,
        inserted: inserted,
        updated: updated,
        rejected: rejected,
        duplicates: duplicates,
        errors: errors.slice(0, 200) // cap payload size; full list is in pallet_uploads_log
      });
    } catch (e) {
      console.error('pallet upload error:', e.message);
      res.status(500).json({ error: 'Upload failed: ' + e.message });
    }
  });

  app.get('/api/pallets/uploads', requireAuth, async function (req, res) {
    try {
      var r = await pool.query('SELECT * FROM pallet_uploads_log ORDER BY created_at DESC LIMIT 50');
      res.json({ uploads: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── FILTER OPTIONS (for dashboard dropdowns) ──

  app.get('/api/pallets/filters/options', requireAuth, async function (req, res) {
    try {
      var cols = ['customer_code', 'customer_name', 'customer_door', 'region', 'salesman_name', 'sales_supervisor', 'brand', 'driver_name', 'transport_type', 'action_owner'];
      var out = {};
      for (var i = 0; i < cols.length; i++) {
        var c = cols[i];
        var r = await pool.query('SELECT DISTINCT ' + c + ' AS v FROM pallet_transactions WHERE ' + c + ' IS NOT NULL AND ' + c + " != '' ORDER BY 1 LIMIT 500");
        out[c] = r.rows.map(function (row) { return row.v; });
      }
      out.collection_status = STATUS_LIST;
      out.ageing_buckets = ['0-7 days', '8-15 days', '16-30 days', '31-60 days', 'Above 60 days'];
      res.json(out);
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── EXECUTIVE KPI DASHBOARD ──

  app.get('/api/pallets/kpis', requireAuth, async function (req, res) {
    try {
      var f = buildFilters(req.query);
      var sql = `
        SELECT
          COALESCE(SUM(t.pallets_delivered),0) AS total_delivered,
          COALESCE(SUM(t.pallets_returned),0) AS total_returned,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS total_outstanding,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value,
          COUNT(DISTINCT CASE WHEN (t.pallets_delivered - t.pallets_returned) > 0 THEN t.customer_code ELSE NULL END) AS customers_outstanding,
          COALESCE(SUM(CASE WHEN (t.pallets_delivered - t.pallets_returned) > 0 AND (CURRENT_DATE - t.delivery_date) > 30 THEN (t.pallets_delivered - t.pallets_returned) ELSE 0 END),0) AS outstanding_above_30,
          COALESCE(SUM(CASE WHEN (t.pallets_delivered - t.pallets_returned) > 0 AND (CURRENT_DATE - t.delivery_date) > 60 THEN (t.pallets_delivered - t.pallets_returned) ELSE 0 END),0) AS outstanding_above_60,
          COUNT(CASE WHEN t.collection_status = 'Collection requested' THEN 1 END) AS collection_requests_pending,
          COUNT(CASE WHEN t.charge_applicable = 'Yes' AND (t.charge_status IS NULL OR t.charge_status <> 'Charged') THEN 1 END) AS charges_pending
        FROM pallet_transactions t
        ${f.where}`;
      var r = await pool.query(sql, f.params);
      var row = r.rows[0];
      var recoveryPct = row.total_delivered > 0 ? +(row.total_returned / row.total_delivered * 100).toFixed(1) : 0;

      // Monthly trend (last 12 months of delivery activity)
      var trendSql = `
        SELECT to_char(date_trunc('month', t.delivery_date), 'YYYY-MM') AS month,
          COALESCE(SUM(t.pallets_delivered),0) AS delivered,
          COALESCE(SUM(t.pallets_returned),0) AS returned,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS outstanding,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value
        FROM pallet_transactions t
        ${f.where}
        GROUP BY 1 ORDER BY 1 DESC LIMIT 12`;
      var trendR = await pool.query(trendSql, f.params);

      res.json({
        total_delivered: +row.total_delivered,
        total_returned: +row.total_returned,
        total_outstanding: +row.total_outstanding,
        recovery_pct: recoveryPct,
        outstanding_value: +row.outstanding_value,
        customers_outstanding: +row.customers_outstanding,
        outstanding_above_30: +row.outstanding_above_30,
        outstanding_above_60: +row.outstanding_above_60,
        collection_requests_pending: +row.collection_requests_pending,
        charges_pending: +row.charges_pending,
        monthly_trend: trendR.rows.reverse()
      });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── AGEING ANALYSIS ──

  app.get('/api/pallets/report/ageing', requireAuth, async function (req, res) {
    try {
      var f = buildFilters(req.query);
      var sql = `
        SELECT ${ageingBucketSQL('(CURRENT_DATE - t.delivery_date)')} AS bucket,
          COUNT(*) AS records,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS outstanding_qty,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value
        FROM pallet_transactions t
        ${f.where ? f.where + ' AND ' : 'WHERE '} (t.pallets_delivered - t.pallets_returned) > 0
        GROUP BY 1`;
      var r = await pool.query(sql, f.params);
      var order = ['0-7 days', '8-15 days', '16-30 days', '31-60 days', 'Above 60 days'];
      var byBucket = {};
      r.rows.forEach(function (row) { byBucket[row.bucket] = row; });
      var result = order.map(function (b) {
        var row = byBucket[b];
        return { bucket: b, records: row ? +row.records : 0, outstanding_qty: row ? +row.outstanding_qty : 0, outstanding_value: row ? +row.outstanding_value : 0 };
      });

      var topOver30 = await pool.query(`
        SELECT t.customer_name, t.customer_code, GREATEST(t.pallets_delivered - t.pallets_returned,0) AS outstanding,
          (CURRENT_DATE - t.delivery_date) AS days, GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL}) AS value_at_risk
        FROM pallet_transactions t ${f.where ? f.where + ' AND ' : 'WHERE '} (t.pallets_delivered - t.pallets_returned) > 0 AND (CURRENT_DATE - t.delivery_date) > 30
        ORDER BY outstanding DESC LIMIT 5`, f.params);
      var topOver60 = await pool.query(`
        SELECT t.customer_name, t.customer_code, GREATEST(t.pallets_delivered - t.pallets_returned,0) AS outstanding,
          (CURRENT_DATE - t.delivery_date) AS days, GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL}) AS value_at_risk
        FROM pallet_transactions t ${f.where ? f.where + ' AND ' : 'WHERE '} (t.pallets_delivered - t.pallets_returned) > 0 AND (CURRENT_DATE - t.delivery_date) > 60
        ORDER BY outstanding DESC LIMIT 5`, f.params);

      res.json({ buckets: result, top_customers_above_30: topOver30.rows, top_customers_above_60: topOver60.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── GENERIC DIMENSION REPORT: customer / door / salesman / region / transport ──

  var DIMENSION_MAP = {
    customer: { col: 't.customer_code', nameCol: 't.customer_name', label: 'Customer' },
    door: { col: 't.customer_door', nameCol: 't.customer_door', label: 'Door / Branch' },
    salesman: { col: 't.salesman_name', nameCol: 't.salesman_name', label: 'Salesman' },
    region: { col: 't.region', nameCol: 't.region', label: 'Region' },
    transport: { col: 't.transport_type', nameCol: 't.transport_type', label: 'Transport Type' },
    site: { col: 't.customer_site_id', nameCol: 't.customer_site_id', label: 'Site ID' }
  };

  app.get('/api/pallets/report/:dimension', requireAuth, async function (req, res) {
    try {
      var dim = DIMENSION_MAP[req.params.dimension];
      if (!dim) return res.status(400).json({ error: 'Unknown report dimension. Use one of: ' + Object.keys(DIMENSION_MAP).join(', ') });
      var f = buildFilters(req.query);
      var sql = `
        SELECT ${dim.nameCol} AS name,
          COUNT(DISTINCT t.customer_code) AS customer_count,
          COALESCE(SUM(t.pallets_delivered),0) AS delivered,
          COALESCE(SUM(t.pallets_returned),0) AS returned,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS outstanding,
          CASE WHEN SUM(t.pallets_delivered) > 0 THEN ROUND((SUM(t.pallets_returned)/SUM(t.pallets_delivered)*100)::numeric,1) ELSE 0 END AS recovery_pct,
          MIN(CASE WHEN (t.pallets_delivered - t.pallets_returned) > 0 THEN t.delivery_date END) AS oldest_outstanding_date,
          COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value,
          COUNT(CASE WHEN t.collection_status NOT IN ('Fully collected','Closed') AND (t.pallets_delivered - t.pallets_returned) > 0 THEN 1 END) AS pending_actions
        FROM pallet_transactions t
        ${f.where}
        WHERE_PLACEHOLDER ${dim.nameCol} IS NOT NULL AND ${dim.nameCol} != ''
        GROUP BY ${dim.nameCol}
        ORDER BY outstanding DESC`;
      sql = sql.replace('WHERE_PLACEHOLDER', f.where ? 'AND' : 'WHERE');
      var r = await pool.query(sql, f.params);
      res.json({ dimension: req.params.dimension, label: dim.label, rows: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── TRANSACTION LIST / DRILL-DOWN ──

  app.get('/api/pallets/transactions', requireAuth, async function (req, res) {
    try {
      var f = buildFilters(req.query);
      var page = Math.max(1, parseInt(req.query.page) || 1);
      var pageSize = Math.min(500, parseInt(req.query.page_size) || 100);
      var offset = (page - 1) * pageSize;

      var countR = await pool.query(`SELECT COUNT(*) AS n FROM pallet_transactions t ${f.where}`, f.params);
      var sql = `
        SELECT t.*, ${CALC_COLUMNS}
        FROM pallet_transactions t
        ${f.where}
        ORDER BY t.delivery_date DESC, t.id DESC
        LIMIT ${pageSize} OFFSET ${offset}`;
      var r = await pool.query(sql, f.params);
      res.json({ total: +countR.rows[0].n, page: page, page_size: pageSize, transactions: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/transactions/:id', requireAuth, async function (req, res) {
    try {
      var r = await pool.query(`SELECT t.*, ${CALC_COLUMNS} FROM pallet_transactions t WHERE t.id=$1`, [req.params.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Record not found' });
      var hist = await pool.query('SELECT * FROM pallet_status_history WHERE transaction_id=$1 ORDER BY created_at DESC', [req.params.id]);
      res.json({ transaction: r.rows[0], status_history: hist.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── EXCEL DOWNLOAD ──

  app.get('/api/pallets/export/:reportType', requireAuth, async function (req, res) {
    try {
      var type = req.params.reportType;
      var f = buildFilters(req.query);
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI';
      wb.created = new Date();
      var titleMap = {
        full: 'AUJAN PALLET — FULL TRANSACTION REPORT',
        customer: 'AUJAN PALLET — CUSTOMER-WISE OUTSTANDING',
        door: 'AUJAN PALLET — DOOR-WISE OUTSTANDING',
        salesman: 'AUJAN PALLET — SALESMAN-WISE REPORT',
        region: 'AUJAN PALLET — REGION-WISE REPORT',
        transport: 'AUJAN PALLET — TRANSPORT COLLECTION REPORT',
        ageing: 'AUJAN PALLET — AGEING REPORT'
      };
      if (!titleMap[type]) return res.status(400).json({ error: 'Unknown report type' });

      var ws = wb.addWorksheet('Report');
      ws.columns = [{ width: 26 }, { width: 22 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 18 }];
      styleTitleRow(ws, titleMap[type], 'G');
      ws.addRow(['Filters', filterSummaryStr(req.query)]);
      ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
      ws.addRow([]);

      if (type === 'full') {
        var txns = await pool.query(`SELECT t.*, ${CALC_COLUMNS} FROM pallet_transactions t ${f.where} ORDER BY t.delivery_date DESC`, f.params);
        ws.columns = [{ width: 12 }, { width: 16 }, { width: 24 }, { width: 14 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }, { width: 20 }, { width: 14 }];
        styleHeaderRow(ws.addRow(['Delivery Date', 'Invoice', 'Customer', 'Site ID', 'Door', 'Delivered', 'Returned', 'Outstanding', 'Recovery %', 'Outstanding Days', 'Status', 'Value at Risk (AED)']));
        var totD = 0, totR = 0, totO = 0, totV = 0;
        txns.rows.forEach(function (t) {
          totD += +t.pallets_delivered; totR += +t.pallets_returned; totO += +t.outstanding_pallets; totV += +t.value_at_risk;
          ws.addRow([t.delivery_date, t.invoice_number, t.customer_name, t.customer_site_id, t.customer_door,
          +t.pallets_delivered, +t.pallets_returned, +t.outstanding_pallets, +t.recovery_pct, +t.outstanding_days, t.collection_status, +t.value_at_risk]);
        });
        styleTotalRow(ws.addRow(['', '', '', '', 'TOTAL', totD, totR, totO, totD > 0 ? +(totR / totD * 100).toFixed(1) : 0, '', '', totV]));
      } else if (type === 'ageing') {
        var ag = await pool.query(`
          SELECT ${ageingBucketSQL('(CURRENT_DATE - t.delivery_date)')} AS bucket, COUNT(*) AS records,
            COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS outstanding_qty,
            COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value
          FROM pallet_transactions t ${f.where ? f.where + ' AND ' : 'WHERE '} (t.pallets_delivered - t.pallets_returned) > 0
          GROUP BY 1`, f.params);
        styleHeaderRow(ws.addRow(['Ageing Bucket', 'Records', 'Outstanding Qty', 'Outstanding Value (AED)']));
        var order = ['0-7 days', '8-15 days', '16-30 days', '31-60 days', 'Above 60 days'];
        var byB = {}; ag.rows.forEach(function (r) { byB[r.bucket] = r; });
        var totQ = 0, totV2 = 0;
        order.forEach(function (b) {
          var row = byB[b] || { records: 0, outstanding_qty: 0, outstanding_value: 0 };
          totQ += +row.outstanding_qty; totV2 += +row.outstanding_value;
          ws.addRow([b, +row.records, +row.outstanding_qty, +row.outstanding_value]);
        });
        styleTotalRow(ws.addRow(['TOTAL', '', totQ, totV2]));
      } else {
        var dim = DIMENSION_MAP[type];
        if (!dim) return res.status(400).json({ error: 'Unknown report type' });
        var sql = `
          SELECT ${dim.nameCol} AS name, COALESCE(SUM(t.pallets_delivered),0) AS delivered, COALESCE(SUM(t.pallets_returned),0) AS returned,
            COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0)),0) AS outstanding,
            CASE WHEN SUM(t.pallets_delivered) > 0 THEN ROUND((SUM(t.pallets_returned)/SUM(t.pallets_delivered)*100)::numeric,1) ELSE 0 END AS recovery_pct,
            COALESCE(SUM(GREATEST(t.pallets_delivered - t.pallets_returned,0) * (${UNIT_COST_SQL})),0) AS outstanding_value
          FROM pallet_transactions t ${f.where}
          ${f.where ? 'AND' : 'WHERE'} ${dim.nameCol} IS NOT NULL AND ${dim.nameCol} != ''
          GROUP BY ${dim.nameCol} ORDER BY outstanding DESC`;
        var rr = await pool.query(sql, f.params);
        styleHeaderRow(ws.addRow([dim.label, 'Delivered', 'Returned', 'Outstanding', 'Recovery %', 'Outstanding Value (AED)']));
        var tD = 0, tR = 0, tO = 0, tV = 0;
        rr.rows.forEach(function (row) {
          tD += +row.delivered; tR += +row.returned; tO += +row.outstanding; tV += +row.outstanding_value;
          ws.addRow([row.name, +row.delivered, +row.returned, +row.outstanding, +row.recovery_pct, +row.outstanding_value]);
        });
        styleTotalRow(ws.addRow(['TOTAL', tD, tR, tO, tD > 0 ? +(tR / tD * 100).toFixed(1) : 0, tV]));
      }

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_Pallet_' + type + '_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('pallet export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });

  // ── STATUS UPDATE (single record — full workflow UI is Phase 2, but the
  // underlying update + history logging is wired now so it's ready) ──

  app.put('/api/pallets/transactions/:id/status', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var b = req.body || {};
      if (!STATUS_LIST.includes(b.status)) return res.status(400).json({ error: 'Invalid status. Must be one of: ' + STATUS_LIST.join(', ') });
      var cur = await pool.query('SELECT collection_status FROM pallet_transactions WHERE id=$1', [req.params.id]);
      if (!cur.rows[0]) return res.status(404).json({ error: 'Record not found' });
      await pool.query('UPDATE pallet_transactions SET collection_status=$1, updated_by=$2, updated_at=NOW() WHERE id=$3', [b.status, req.user.uid, req.params.id]);
      await pool.query(
        `INSERT INTO pallet_status_history (transaction_id, previous_status, new_status, updated_by, comment, next_action_date)
         VALUES ($1,$2,$3,$4,$5,$6)`,
        [req.params.id, cur.rows[0].collection_status, b.status, req.user.username, b.comment || null, b.next_action_date || null]
      );
      res.json({ ok: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ============================================================
  // AUJAN SUPPLY TRACKING — Inbound (received from Aujan) and
  // Return to Aujan. Warehouse/shipment-level, kept deliberately
  // separate from the customer delivery/recovery flow above.
  // ============================================================

  var SUPPLY_LABELS = { inbound: 'Inbound — Received from Aujan', 'return': 'Return to Aujan' };
  var SUPPLY_HEADER_MAP = [
    ['Date', 'entry_date', 'date'],
    ['Reference / Shipment No', 'reference', 'text'],
    ['Number of Trucks', 'truck_count', 'number'],
    ['Pallet Type', 'pallet_type', 'text'],
    ['Pallet Qty', 'pallet_qty', 'number'],
    ['Warehouse', 'warehouse', 'text'],
    ['Remarks', 'remarks', 'text']
  ];

  function validSupplyType(t) { return t === 'inbound' || t === 'return'; }

  app.get('/api/pallets/supply/upload-template/:type', requireAuth, async function (req, res) {
    try {
      var type = req.params.type;
      if (!validSupplyType(type)) return res.status(400).json({ error: 'type must be inbound or return' });
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();
      var ws = wb.addWorksheet('Data');
      ws.columns = SUPPLY_HEADER_MAP.map(function () { return { width: 22 }; });
      styleHeaderRow(ws.addRow(SUPPLY_HEADER_MAP.map(function (h) { return h[0]; })));
      ws.addRow(['Mandatory: Date, Pallet Qty. One row per shipment/day — e.g. "5 trucks, 200 pallets" is one row.']);
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_' + (type === 'inbound' ? 'Inbound' : 'Return') + '_Template.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/pallets/supply/upload/:type', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      var type = req.params.type;
      if (!validSupplyType(type)) return res.status(400).json({ error: 'type must be inbound or return' });
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
      catch (e) { return res.status(400).json({ error: 'Could not read that file. Upload the .xlsx template or a CSV export of it.' }); }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      var rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rawRows.length) return res.status(400).json({ error: 'The file has no data rows.' });

      var batchId = 'SUP-' + type.toUpperCase() + '-' + Date.now();
      var inserted = 0, rejected = 0;
      var errors = [];

      for (var i = 0; i < rawRows.length; i++) {
        var raw = rawRows[i], rowNum = i + 2, rec = {};
        SUPPLY_HEADER_MAP.forEach(function (h) {
          var display = h[0], field = h[1], t = h[2];
          var val = raw[display];
          if (val === undefined) {
            var key = Object.keys(raw).find(function (k) { return k.trim().toLowerCase() === display.trim().toLowerCase(); });
            if (key) val = raw[key];
          }
          if (t === 'date') rec[field] = toDateStr(val);
          else if (t === 'number') rec[field] = toNum(val);
          else rec[field] = (val === undefined || val === null) ? null : String(val).trim();
        });

        if (!rec.entry_date || !rec.pallet_qty) {
          rejected++;
          errors.push({ row: rowNum, reason: 'Missing mandatory field(s): ' + (!rec.entry_date ? 'Date ' : '') + (!rec.pallet_qty ? 'Pallet Qty' : '') });
          continue;
        }
        rec.pallet_type = rec.pallet_type || 'Standard';

        try {
          var ins = await pool.query(
            `INSERT INTO aujan_supply_log (entry_type, entry_date, reference, truck_count, pallet_type, pallet_qty, warehouse, remarks, upload_batch_id, created_by, created_by_name)
             VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) RETURNING id`,
            [type, rec.entry_date, rec.reference, rec.truck_count || null, rec.pallet_type, rec.pallet_qty, rec.warehouse, rec.remarks, batchId, req.user.uid, req.user.full_name || req.user.username]
          );
          inserted++;
        } catch (rowErr) {
          rejected++;
          errors.push({ row: rowNum, reason: rowErr.message });
        }
      }

      await auditLog(req.user.uid, req.user.username, 'PALLET_SUPPLY_UPLOAD', type + ': ' + req.file.originalname + ' — inserted ' + inserted + ', rejected ' + rejected, req.headers['x-forwarded-for'] || '');
      res.json({ batch_id: batchId, type: type, total_rows: rawRows.length, inserted: inserted, rejected: rejected, errors: errors.slice(0, 200) });
    } catch (e) {
      console.error('supply upload error:', e.message);
      res.status(500).json({ error: 'Upload failed: ' + e.message });
    }
  });

  app.get('/api/pallets/supply/summary', requireAuth, async function (req, res) {
    try {
      var params = [];
      var dateClauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); dateClauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); dateClauses.push('entry_date <= $' + params.length); }
      var dateWhere = dateClauses.length ? (' AND ' + dateClauses.join(' AND ')) : '';

      var recv = await pool.query(`SELECT COALESCE(SUM(pallet_qty),0) AS qty, COALESCE(SUM(truck_count),0) AS trucks, COUNT(*) AS entries FROM aujan_supply_log WHERE entry_type='inbound'${dateWhere}`, params);
      var ret = await pool.query(`SELECT COALESCE(SUM(pallet_qty),0) AS qty, COALESCE(SUM(truck_count),0) AS trucks, COUNT(*) AS entries FROM aujan_supply_log WHERE entry_type='return'${dateWhere}`, params);
      var totalReceived = +recv.rows[0].qty, totalReturned = +ret.rows[0].qty;

      res.json({
        total_received: totalReceived,
        total_returned_to_aujan: totalReturned,
        balance_with_aki: totalReceived - totalReturned,
        inbound_trucks: +recv.rows[0].trucks,
        inbound_entries: +recv.rows[0].entries,
        return_trucks: +ret.rows[0].trucks,
        return_entries: +ret.rows[0].entries
      });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/supply/log', requireAuth, async function (req, res) {
    try {
      var type = req.query.type;
      var clauses = [], params = [];
      if (type && validSupplyType(type)) { params.push(type); clauses.push('entry_type = $' + params.length); }
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var page = Math.max(1, parseInt(req.query.page) || 1);
      var pageSize = Math.min(500, parseInt(req.query.page_size) || 100);
      var offset = (page - 1) * pageSize;
      var countR = await pool.query('SELECT COUNT(*) AS n FROM aujan_supply_log ' + where, params);
      var r = await pool.query('SELECT * FROM aujan_supply_log ' + where + ' ORDER BY entry_date DESC, id DESC LIMIT ' + pageSize + ' OFFSET ' + offset, params);
      res.json({ total: +countR.rows[0].n, page: page, page_size: pageSize, entries: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/supply/export/:type', requireAuth, async function (req, res) {
    try {
      var type = req.params.type;
      if (!validSupplyType(type)) return res.status(400).json({ error: 'type must be inbound or return' });
      var params = [type], clauses = ['entry_type = $1'];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var r = await pool.query('SELECT * FROM aujan_supply_log WHERE ' + clauses.join(' AND ') + ' ORDER BY entry_date DESC', params);

      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();
      var ws = wb.addWorksheet('Report');
      ws.columns = [{ width: 14 }, { width: 22 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 30 }];
      styleTitleRow(ws, 'AUJAN ' + (type === 'inbound' ? 'INBOUND RECEIVING' : 'RETURN TO AUJAN') + ' REPORT', 'F');
      ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
      ws.addRow([]);
      styleHeaderRow(ws.addRow(['Date', 'Reference', 'Trucks', 'Pallet Type', 'Pallet Qty', 'Remarks']));
      var totQty = 0, totTrucks = 0;
      r.rows.forEach(function (row) {
        totQty += +row.pallet_qty; totTrucks += +(row.truck_count || 0);
        ws.addRow([row.entry_date, row.reference, +row.truck_count || 0, row.pallet_type, +row.pallet_qty, row.remarks]);
      });
      styleTotalRow(ws.addRow(['', 'TOTAL', totTrucks, '', totQty, '']));

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_' + type + '_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
  });

  app.get('/api/pallets/meta', requireAuth, function (req, res) {
    res.json({ statuses: STATUS_LIST, departments: DEPARTMENTS, priorities: PRIORITIES });
  });
};
