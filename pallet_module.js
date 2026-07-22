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

  const GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', REDFLAG = 'FFFDE0DE', GREENFLAG = 'FFE1F5E9', REDTEXT = 'FFB33A3A', GREENTEXT = 'FF2E7D4F';

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
    ['Collected (Yes/No)', 'collected_flag', 'text'],
    ['Reason Not Collected', 'reason_not_collected', 'text'],
    ['Next Action Date', 'next_action_date', 'date'],
    ['Action Owner', 'action_owner', 'text'],
    ['Charge Applicable (Yes/No)', 'charge_applicable', 'text'],
    ['Charge Amount', 'charge_amount', 'number'],
    ['Charge Status', 'charge_status', 'text'],
    ['Remarks', 'remarks', 'text']
  ];
  const MANDATORY_FIELDS = ['delivery_date', 'invoice_number', 'customer_site_id', 'customer_name', 'pallets_delivered'];

  // Fuzzy keyword matching for the customer-facing upload — covers both this
  // app's own template headers AND the real daily "Order Details" file the
  // transport team actually sends (DATE, ORDER CODE, PALLET COUNT, CITY,
  // LOCATION_ID, CUSTOMER, CUSTOMER ADDRESS, VEHICLE_ID, DRIVER_ID, etc — no
  // Door field, no Customer Code column). Order matters: more specific
  // patterns are checked before generic fallbacks.
  function normalizeTxnHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
  function matchTransactionField(header) {
    var h = normalizeTxnHeader(header);
    if (h.indexOf('returndate') !== -1) return ['return_date', 'date'];
    if (h.indexOf('collectionrequest') !== -1 && h.indexOf('date') !== -1) return ['collection_request_date', 'date'];
    if (h.indexOf('collectionattempt') !== -1 && h.indexOf('date') !== -1) return ['collection_attempt_date', 'date'];
    if (h.indexOf('nextaction') !== -1 && h.indexOf('date') !== -1) return ['next_action_date', 'date'];
    if (h.indexOf('date') !== -1) return ['delivery_date', 'date'];

    if (h.indexOf('invoice') !== -1) return ['invoice_number', 'text'];
    if (h.indexOf('ordercode') !== -1) return ['invoice_number', 'text']; // real file's order identifier plays the same role
    if (h.indexOf('salesorder') !== -1) return ['sales_order_number', 'text'];

    if (h.indexOf('customercode') !== -1) return ['customer_code', 'text'];
    if (h.indexOf('customeraddress') !== -1) return ['ship_to_location', 'text'];
    if (h.indexOf('customerconfirmation') !== -1) return ['customer_confirmation', 'text'];
    if (h.indexOf('customername') !== -1 || h === 'customer') return ['customer_name', 'text'];

    if (h.indexOf('siteid') !== -1 || h.indexOf('locationid') !== -1) return ['customer_site_id', 'text'];
    if (h.indexOf('shipto') !== -1) return ['ship_to_location', 'text'];
    if (h.indexOf('door') !== -1) return ['customer_door', 'text'];
    if (h.indexOf('region') !== -1 || h.indexOf('emirate') !== -1 || h === 'city') return ['region', 'text'];

    if (h.indexOf('salesman') !== -1 || h.indexOf('salesperson') !== -1) return ['salesman_name', 'text'];
    if (h.indexOf('supervisor') !== -1) return ['sales_supervisor', 'text'];
    if (h.indexOf('brand') !== -1) return ['brand', 'text'];
    if (h.indexOf('pallettype') !== -1) return ['pallet_type', 'text'];

    if (h.indexOf('palletcount') !== -1 || (h.indexOf('pallet') !== -1 && (h.indexOf('deliver') !== -1 || h.indexOf('dispatch') !== -1))) return ['pallets_delivered', 'number'];
    if (h.indexOf('pallet') !== -1 && h.indexOf('return') !== -1) return ['pallets_returned', 'number'];

    if (h.indexOf('transporttype') !== -1) return ['transport_type', 'text'];
    if (h.indexOf('vehicle') !== -1) return ['vehicle_number', 'text'];
    if (h.indexOf('driverid') !== -1 || h === 'driver' || h.indexOf('drivername') !== -1) return ['driver_name', 'text'];

    if (h.indexOf('collectionstatus') !== -1) return ['collection_status', 'text'];
    if (h.indexOf('reasonnotcollected') !== -1) return ['reason_not_collected', 'text'];
    // Simple Yes/No "Collected or not" column — used to derive pallets_returned
    // and collection_status below, not stored as its own DB column.
    if (h.indexOf('collected') !== -1) return ['collected_flag', 'text'];
    if (h.indexOf('actionowner') !== -1) return ['action_owner', 'text'];
    if (h.indexOf('chargeapplicable') !== -1) return ['charge_applicable', 'text'];
    if (h.indexOf('chargeamount') !== -1) return ['charge_amount', 'number'];
    if (h.indexOf('chargestatus') !== -1) return ['charge_status', 'text'];
    if (h.indexOf('specialinstruction') !== -1 || h.indexOf('remark') !== -1) return ['remarks', 'text'];
    return null;
  }

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
      // Migration: real-world transport data (e.g. the daily Order Details email
      // from the transport team) has no Door field at all, so it can no longer be
      // a hard requirement. Relax it on already-deployed databases too.
      await pool.query(`ALTER TABLE pallet_transactions ALTER COLUMN customer_door DROP NOT NULL`);
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

      // ── Aujan Pallet Inbound & Outbound Tracker (warehouse-level running
      // ledger, matches the transport team's actual Excel). One row per
      // truck/tour/invoice. Opening/Closing balance are ALWAYS recomputed
      // server-side (window function, ordered by date then row sequence)
      // rather than trusted from the upload — this guarantees the running
      // balance can never drift from a copy-paste error in someone's Excel.
      await pool.query(`CREATE TABLE IF NOT EXISTS aujan_ledger (
        id SERIAL PRIMARY KEY,
        entry_date DATE NOT NULL,
        truck_ref TEXT NOT NULL,
        inbound_pallets NUMERIC NOT NULL DEFAULT 0,
        outbound_pallets NUMERIC NOT NULL DEFAULT 0,
        return_received_from_delivery NUMERIC NOT NULL DEFAULT 0,
        damage_pallets NUMERIC NOT NULL DEFAULT 0,
        pallets_returned_to_aujan NUMERIC NOT NULL DEFAULT 0,
        pallet_retained_by_customer NUMERIC NOT NULL DEFAULT 0,
        reported_opening_pallets NUMERIC,
        reported_closing_balance NUMERIC,
        remarks TEXT,
        upload_batch_id TEXT,
        created_by INT REFERENCES users(id),
        created_by_name TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      // Migration: earlier version of this module wrongly made truck_ref alone
      // UNIQUE. Tour numbers legitimately recur across different dates, so that
      // constraint caused later uploads to silently overwrite earlier rows
      // instead of adding new ones. Drop it if it exists on an already-deployed DB.
      await pool.query(`ALTER TABLE aujan_ledger DROP CONSTRAINT IF EXISTS aujan_ledger_truck_ref_key`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_ledger_date ON aujan_ledger(entry_date, id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_ledger_truck_date ON aujan_ledger(entry_date, truck_ref)`);

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
    var str = String(val).trim();
    // Explicit DD/MM/YYYY or DD-MM-YYYY first — this is how UAE-format text
    // dates actually show up in real uploaded files (a cell stored as text,
    // not a real Excel date type). JS's bare `new Date(string)` assumes
    // MM/DD/YYYY and silently returns Invalid Date for e.g. "14/07/2026"
    // since there's no 14th month — that failure must be caught here, not
    // left to fall through to the ambiguous native parser below.
    var dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (dmy) {
      var day = +dmy[1], month = +dmy[2], year = +dmy[3];
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        var dd = new Date(Date.UTC(year, month - 1, day));
        if (!isNaN(dd.getTime())) return dd.toISOString().slice(0, 10);
      }
    }
    var parsed = new Date(str);
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
      var note = ws.addRow(['Mandatory: Delivery Date, Invoice Number, Customer Site ID (or Location ID), Customer Name, Pallets Delivered. Customer Code is auto-derived from Site ID if not provided. Door is optional. Dates as DD-MM-YYYY or YYYY-MM-DD. Outstanding pallets, Recovery %, Ageing and Value at Risk are calculated automatically — do not enter them.']);
      ws.mergeCells('A' + (headerRow.number) + ':' + 'A' + headerRow.number); // no-op, keeps layout stable
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_Pallet_Upload_Template.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── UPLOAD DATA (Excel or CSV) ──

  // Clears the main Pallet Recovery Dashboard dataset (pallet_transactions) —
  // separate from /api/pallets/ledger/clear, which only clears the Inbound &
  // Outbound Tracker ledger below it on the page.
  app.delete('/api/pallets/clear', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM pallet_transactions');
      await auditLog(req.user.uid, req.user.username, 'PALLET_DATA_CLEAR', 'Cleared all pallet transactions (' + r.rowCount + ')', req.headers['x-forwarded-for'] || '');
      res.json({ ok: true, deleted: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/pallets/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try {
        wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
      } catch (parseErr) {
        return res.status(400).json({ error: 'Could not read that file. Upload the .xlsx template or a CSV export of it.' });
      }

      var batchId = 'PLT-' + Date.now();
      var inserted = 0, updated = 0, rejected = 0, duplicates = 0;
      var errors = [];
      var seenInBatch = {}; // key -> row# for in-file duplicate detection

      // Pick the sheet that actually looks like transaction data — the transport
      // team's real email has TWO sheets (a daily-totals summary and an
      // "ORDER DETAILS" sheet with the real per-customer rows), and the one we
      // want isn't always first. Score every sheet's header row and take the
      // best match, requiring at minimum a customer-name-ish and a
      // pallets-delivered-ish column before considering it a candidate.
      var bestSheetName = wb.SheetNames[0], bestScore = -1;
      wb.SheetNames.forEach(function (name) {
        var s = wb.Sheets[name];
        var firstRow = XLSX.utils.sheet_to_json(s, { header: 1, defval: '' })[0] || [];
        var score = 0, hasCustomer = false, hasQty = false;
        firstRow.forEach(function (cell) {
          var m = matchTransactionField(cell);
          if (m) { score++; if (m[0] === 'customer_name') hasCustomer = true; if (m[0] === 'pallets_delivered') hasQty = true; }
        });
        if (hasCustomer && hasQty && score > bestScore) { bestScore = score; bestSheetName = name; }
      });

      var sheet = wb.Sheets[bestSheetName];
      var rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rawRows.length) return res.status(400).json({ error: 'The file has no data rows.' });

      var sampleHeaders = Object.keys(rawRows[0]);
      var fieldByHeader = {};
      sampleHeaders.forEach(function (h) { var m = matchTransactionField(h); if (m) fieldByHeader[h] = m; });

      var batchId = 'PLT-' + Date.now();
      var inserted = 0, updated = 0, rejected = 0, duplicates = 0;
      var errors = [];
      var seenInBatch = {}; // key -> row# for in-file duplicate detection

      for (var i = 0; i < rawRows.length; i++) {
        var raw = rawRows[i];
        var rowNum = i + 2; // header is row 1
        var rec = {};
        Object.keys(raw).forEach(function (h) {
          var m = fieldByHeader[h];
          if (!m) return;
          var field = m[0], type = m[1], val = raw[h];
          if (type === 'date') rec[field] = toDateStr(val);
          else if (type === 'number') rec[field] = toNum(val);
          else rec[field] = (val === undefined || val === null) ? null : String(val).trim();
        });
        // Location ID / Site ID doubles as Customer Code when the file has no
        // separate code column (true for the real transport-team file).
        if (!rec.customer_code && rec.customer_site_id) rec.customer_code = rec.customer_site_id;

        // A simple "Collected or not" Yes/No column (added on request) drives
        // pallets_returned and collection_status directly — this is the piece
        // that gives per-customer return visibility instead of everything
        // sitting at 0 returned forever.
        if (rec.collected_flag !== undefined && rec.collected_flag !== null && rec.collected_flag !== '') {
          var flagNorm = String(rec.collected_flag).trim().toLowerCase();
          var isCollected = ['yes', 'y', '1', 'true', 'collected'].indexOf(flagNorm) !== -1;
          var isNotCollected = ['no', 'n', '0', 'false', 'not collected', 'notcollected', 'pending'].indexOf(flagNorm) !== -1;
          if (isCollected) {
            rec.pallets_returned = rec.pallets_delivered;
            rec.collection_status = 'Fully collected';
            if (!rec.return_date) rec.return_date = rec.delivery_date;
          } else if (isNotCollected) {
            rec.pallets_returned = 0;
            rec.collection_status = 'Pending customer confirmation';
          }
          // Anything else unrecognized (blank, typo) — leave pallets_returned/
          // collection_status to the normal defaults rather than guessing.
        }
        // Real delivery-only files (no return tracking yet) legitimately have
        // no Pallets Returned column at all — default to 0 rather than reject.
        if (rec.pallets_returned === undefined || rec.pallets_returned === null || rec.pallets_returned === '') rec.pallets_returned = 0;

        // Mandatory field validation
        var missing = MANDATORY_FIELDS.filter(function (f) {
          return rec[f] === null || rec[f] === undefined || rec[f] === '';
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
  // AUJAN PALLET INBOUND & OUTBOUND TRACKER — the real warehouse ledger
  // (one row per truck/tour/invoice), kept deliberately separate from the
  // customer delivery/recovery flow above. Opening/Closing balance are
  // always computed server-side via a running window function ordered by
  // (date, id) — never trusted from the upload — so the ledger can't drift
  // from an Excel formula error. The row's own reported values are still
  // stored for reference and cross-checked (flagged if they disagree).
  // ============================================================

  // Column headers in real-world files vary (typos, spacing, "Pallets" vs
  // "Palets"). Match by keyword rather than exact text so re-uploads of the
  // transport team's actual file work without reformatting.
  function normalizeHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
  function matchLedgerField(header) {
    var h = normalizeHeader(header);
    if (h.indexOf('date') !== -1) return ['entry_date', 'date'];
    if (h.indexOf('truck') !== -1 || h.indexOf('tour') !== -1 || h.indexOf('invoice') !== -1) return ['truck_ref', 'text'];
    if (h.indexOf('opening') !== -1) return ['reported_opening_pallets', 'number'];
    if (h.indexOf('inbound') !== -1) return ['inbound_pallets', 'number'];
    if (h.indexOf('outbound') !== -1) return ['outbound_pallets', 'number'];
    if (h.indexOf('return') !== -1 && h.indexOf('deliver') !== -1) return ['return_received_from_delivery', 'number'];
    if (h.indexOf('damage') !== -1) return ['damage_pallets', 'number'];
    if (h.indexOf('return') !== -1 && h.indexOf('aujan') !== -1) return ['pallets_returned_to_aujan', 'number'];
    if (h.indexOf('closing') !== -1) return ['reported_closing_balance', 'number'];
    if (h.indexOf('retain') !== -1) return ['pallet_retained_by_customer', 'number'];
    if (h.indexOf('remark') !== -1) return ['remarks', 'text'];
    return null;
  }

  var LEDGER_TEMPLATE_HEADERS = ['Date', 'Truck/Tour/Invoice #', 'Opening Pallets', 'Inbound Pallets', 'Outbound Pallets',
    'Return Received from Delivery', 'Damage Pallets', 'Pallets Returned to Aujan', 'Closing Balance', 'Pallet Retained by Customer', 'Remarks'];

  // Reusable running-balance SQL. net_movement = inbound - outbound + return_received - damage - returned_to_aujan.
  // computed_closing = cumulative net movement through this row. computed_opening = cumulative through the PRIOR row.
  var LEDGER_CALC_SQL = `
    l.*,
    (l.inbound_pallets - l.outbound_pallets + l.return_received_from_delivery - l.damage_pallets - l.pallets_returned_to_aujan) AS net_movement,
    COALESCE(SUM(l.inbound_pallets - l.outbound_pallets + l.return_received_from_delivery - l.damage_pallets - l.pallets_returned_to_aujan)
      OVER (ORDER BY l.entry_date, l.id ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING), 0) AS computed_opening_balance,
    SUM(l.inbound_pallets - l.outbound_pallets + l.return_received_from_delivery - l.damage_pallets - l.pallets_returned_to_aujan)
      OVER (ORDER BY l.entry_date, l.id ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS computed_closing_balance`;

  // Superadmin-only escape hatch: wipe all ledger rows for a clean re-import.
  // Needed once here because an earlier version of the dedup logic could
  // silently collapse legitimate same-tour-different-date rows — anyone
  // who uploaded under that version should clear and re-upload the full
  // file fresh rather than trying to patch the data in place.
  app.delete('/api/pallets/ledger/clear', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM aujan_ledger');
      await auditLog(req.user.uid, req.user.username, 'PALLET_LEDGER_CLEAR', 'Cleared all ledger rows (' + r.rowCount + ')', req.headers['x-forwarded-for'] || '');
      res.json({ ok: true, deleted: r.rowCount });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/ledger/upload-template', requireAuth, async function (req, res) {
    try {
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();
      var ws = wb.addWorksheet('Ledger');
      ws.columns = LEDGER_TEMPLATE_HEADERS.map(function () { return { width: 20 }; });
      styleHeaderRow(ws.addRow(LEDGER_TEMPLATE_HEADERS));
      ws.addRow(['Mandatory: Date, Truck/Tour/Invoice #. Opening Pallets and Closing Balance are recalculated automatically from the movement columns — enter them if you have them, but they will not override the running balance.']);
      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_Pallet_Ledger_Template.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/pallets/ledger/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      var wb;
      try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
      catch (e) { return res.status(400).json({ error: 'Could not read that file. Upload the .xlsx template or a CSV export of it.' }); }
      var sheet = wb.Sheets[wb.SheetNames[0]];
      // Read as a raw grid first — don't assume row 1 is the header row. Real-world
      // files (like this one) often have a summary block sitting above the actual
      // table, so we scan for the row that actually looks like ledger headers.
      var grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!grid.length) return res.status(400).json({ error: 'The file has no data rows.' });

      var headerRowIdx = -1, bestScore = -1;
      for (var ri = 0; ri < Math.min(grid.length, 40); ri++) {
        var row = grid[ri];
        var hasDate = false, hasTruck = false, score = 0;
        row.forEach(function (cell) {
          var m = matchLedgerField(cell);
          if (m) { score++; if (m[0] === 'entry_date') hasDate = true; if (m[0] === 'truck_ref') hasTruck = true; }
        });
        if (hasDate && hasTruck && score > bestScore) { bestScore = score; headerRowIdx = ri; }
      }
      if (headerRowIdx === -1) {
        return res.status(400).json({ error: 'Could not find a header row with both a Date column and a Truck/Tour/Invoice # column anywhere in the first 40 rows of the file. If your sheet has a summary block above the table, that\'s fine — just make sure the table\'s own header row (Date, Truck/Tour/Invoice #, Inbound Pallets, etc.) is somewhere below it.' });
      }

      var headerRow = grid[headerRowIdx];
      var fieldByIndex = {};
      headerRow.forEach(function (h, idx) { var m = matchLedgerField(h); if (m) fieldByIndex[idx] = m; });
      var dataRows = grid.slice(headerRowIdx + 1);

      var batchId = 'LEDGER-' + Date.now();
      var inserted = 0, updated = 0, rejected = 0, totalDataRows = 0;
      var errors = [];
      var occurrenceCount = {}; // tracks how many times (date|truck_ref) has appeared so far in THIS upload

      for (var i = 0; i < dataRows.length; i++) {
        var rowArr = dataRows[i];
        var rowNum = headerRowIdx + i + 2; // grid is 0-indexed; Excel rows are 1-indexed, and dataRows starts right after the header row
        if (!rowArr || rowArr.every(function (c) { return c === '' || c === undefined || c === null; })) continue; // skip fully blank rows (e.g. spacer rows)
        totalDataRows++;

        var rec = {};
        Object.keys(fieldByIndex).forEach(function (idxStr) {
          var idx = +idxStr, m = fieldByIndex[idx], field = m[0], type = m[1], val = rowArr[idx];
          if (type === 'date') rec[field] = toDateStr(val);
          else if (type === 'number') rec[field] = toNum(val);
          else rec[field] = (val === undefined || val === null || val === '') ? null : String(val).trim();
        });
        ['inbound_pallets', 'outbound_pallets', 'return_received_from_delivery', 'damage_pallets', 'pallets_returned_to_aujan', 'pallet_retained_by_customer'].forEach(function (f) {
          if (rec[f] === undefined) rec[f] = 0;
        });

        if (!rec.entry_date || !rec.truck_ref) {
          rejected++;
          errors.push({ row: rowNum, reason: 'Missing mandatory field(s): ' + (!rec.entry_date ? 'Date ' : '') + (!rec.truck_ref ? 'Truck/Tour/Invoice #' : '') });
          continue;
        }

        try {
          // Tour/truck references legitimately recur across different dates (a route
          // that runs daily), so the natural key is (date, truck_ref) — and even that
          // can repeat within a single day if the same tour runs twice. Track how many
          // times we've seen this exact (date, truck_ref) pair so far in this upload,
          // and match it against the Nth existing DB row for that pair (ordered by id).
          // This keeps re-uploads idempotent without ever collapsing distinct trips.
          var dedupKey = rec.entry_date + '|' + rec.truck_ref;
          occurrenceCount[dedupKey] = (occurrenceCount[dedupKey] || 0) + 1;
          var occurrence = occurrenceCount[dedupKey];

          var existing = await pool.query('SELECT id FROM aujan_ledger WHERE entry_date=$1 AND truck_ref=$2 ORDER BY id ASC', [rec.entry_date, rec.truck_ref]);
          if (existing.rows.length >= occurrence) {
            await pool.query(
              `UPDATE aujan_ledger SET inbound_pallets=$1, outbound_pallets=$2, return_received_from_delivery=$3,
                damage_pallets=$4, pallets_returned_to_aujan=$5, pallet_retained_by_customer=$6,
                reported_opening_pallets=$7, reported_closing_balance=$8, remarks=$9, upload_batch_id=$10
               WHERE id=$11`,
              [rec.inbound_pallets, rec.outbound_pallets, rec.return_received_from_delivery,
              rec.damage_pallets, rec.pallets_returned_to_aujan, rec.pallet_retained_by_customer,
              rec.reported_opening_pallets || null, rec.reported_closing_balance || null, rec.remarks, batchId, existing.rows[occurrence - 1].id]
            );
            updated++;
          } else {
            await pool.query(
              `INSERT INTO aujan_ledger (entry_date, truck_ref, inbound_pallets, outbound_pallets, return_received_from_delivery,
                damage_pallets, pallets_returned_to_aujan, pallet_retained_by_customer, reported_opening_pallets, reported_closing_balance,
                remarks, upload_batch_id, created_by, created_by_name)
               VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`,
              [rec.entry_date, rec.truck_ref, rec.inbound_pallets, rec.outbound_pallets, rec.return_received_from_delivery,
              rec.damage_pallets, rec.pallets_returned_to_aujan, rec.pallet_retained_by_customer,
              rec.reported_opening_pallets || null, rec.reported_closing_balance || null, rec.remarks, batchId, req.user.uid, req.user.full_name || req.user.username]
            );
            inserted++;
          }
        } catch (rowErr) {
          rejected++;
          errors.push({ row: rowNum, reason: rowErr.message });
        }
      }

      await auditLog(req.user.uid, req.user.username, 'PALLET_LEDGER_UPLOAD', req.file.originalname + ' — inserted ' + inserted + ', updated ' + updated + ', rejected ' + rejected, req.headers['x-forwarded-for'] || '');
      res.json({ batch_id: batchId, total_rows: totalDataRows, inserted: inserted, updated: updated, rejected: rejected, errors: errors.slice(0, 200) });
    } catch (e) {
      console.error('ledger upload error:', e.message);
      res.status(500).json({ error: 'Upload failed: ' + e.message });
    }
  });

  app.get('/api/pallets/ledger/summary', requireAuth, async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';

      var sumR = await pool.query(`
        SELECT COALESCE(SUM(inbound_pallets),0) AS inbound, COALESCE(SUM(outbound_pallets),0) AS outbound,
          COALESCE(SUM(return_received_from_delivery),0) AS return_received, COALESCE(SUM(damage_pallets),0) AS damage,
          COALESCE(SUM(pallets_returned_to_aujan),0) AS returned_to_aujan, COUNT(*) AS entries
        FROM aujan_ledger ${where}`, params);
      var s = sumR.rows[0];

      // Current warehouse balance = latest row's computed running closing balance (unfiltered — this is a point-in-time truth, not a date-range sum).
      var balR = await pool.query(`
        SELECT SUM(inbound_pallets - outbound_pallets + return_received_from_delivery - damage_pallets - pallets_returned_to_aujan) AS balance
        FROM aujan_ledger`);
      var currentBalance = balR.rows[0].balance !== null ? +balR.rows[0].balance : 0;

      res.json({
        inbound_pallets: +s.inbound,
        outbound_pallets: +s.outbound,
        return_received_from_delivery: +s.return_received,
        damage_pallets: +s.damage,
        pallets_returned_to_aujan: +s.returned_to_aujan,
        customer_retained: +s.return_received - +s.outbound,
        current_warehouse_balance: currentBalance,
        entries: +s.entries
      });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/ledger/entries', requireAuth, async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('l.entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('l.entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var page = Math.max(1, parseInt(req.query.page) || 1);
      var pageSize = Math.min(500, parseInt(req.query.page_size) || 100);
      var offset = (page - 1) * pageSize;

      var countR = await pool.query('SELECT COUNT(*) AS n FROM aujan_ledger l ' + where, params);
      // Window function needs the full ordered set to compute a correct running balance, so we compute
      // over everything matching the filter, then page the result in JS (page sizes here are small).
      var full = await pool.query(`SELECT ${LEDGER_CALC_SQL} FROM aujan_ledger l ${where} ORDER BY l.entry_date, l.id`, params);
      var rows = full.rows.slice().reverse(); // newest first for display
      var pageRows = rows.slice(offset, offset + pageSize);
      res.json({ total: +countR.rows[0].n, page: page, page_size: pageSize, entries: pageRows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/ledger/trend', requireAuth, async function (req, res) {
    try {
      var period = req.query.period;
      if (!['day', 'week', 'month'].includes(period)) period = 'day';
      var limit = { day: 30, week: 16, month: 12 }[period];
      var params = [period], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var sql = `
        SELECT date_trunc($1, entry_date) AS period_start,
          COALESCE(SUM(inbound_pallets),0) AS inbound, COALESCE(SUM(outbound_pallets),0) AS outbound,
          COALESCE(SUM(return_received_from_delivery),0) AS return_received, COALESCE(SUM(damage_pallets),0) AS damage,
          COALESCE(SUM(pallets_returned_to_aujan),0) AS returned_to_aujan
        FROM aujan_ledger ${where}
        GROUP BY 1 ORDER BY 1 DESC LIMIT ${limit}`;
      var r = await pool.query(sql, params);
      var points = r.rows.reverse().map(function (row) {
        return { period_start: row.period_start, inbound: +row.inbound, outbound: +row.outbound, return_received: +row.return_received, damage: +row.damage, returned_to_aujan: +row.returned_to_aujan };
      });
      res.json({ period: period, points: points });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.get('/api/pallets/ledger/export', requireAuth, async function (req, res) {
    try {
      var params = [], clauses = [];
      if (req.query.date_from) { params.push(req.query.date_from); clauses.push('l.entry_date >= $' + params.length); }
      if (req.query.date_to) { params.push(req.query.date_to); clauses.push('l.entry_date <= $' + params.length); }
      var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
      var r = await pool.query(`SELECT ${LEDGER_CALC_SQL} FROM aujan_ledger l ${where} ORDER BY l.entry_date, l.id`, params);
      var rows = r.rows;

      // Aggregate totals (same numbers the dashboard KPI cards show).
      var totals = rows.reduce(function (a, row) {
        a.inbound += +row.inbound_pallets; a.outbound += +row.outbound_pallets;
        a.returnRecv += +row.return_received_from_delivery; a.damage += +row.damage_pallets;
        a.returnedAujan += +row.pallets_returned_to_aujan;
        return a;
      }, { inbound: 0, outbound: 0, returnRecv: 0, damage: 0, returnedAujan: 0 });
      var customerRetained = totals.returnRecv - totals.outbound;
      var currentBalance = rows.length ? +rows[rows.length - 1].computed_closing_balance : 0;

      // Pivot-style breakdowns (month and week), reusing the same grouping the trend chart uses.
      function groupBy(period) {
        var buckets = {};
        rows.forEach(function (row) {
          var d = new Date(row.entry_date);
          var key;
          if (period === 'month') key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
          else {
            var weekStart = new Date(d); weekStart.setDate(d.getDate() - d.getDay());
            key = weekStart.toISOString().slice(0, 10);
          }
          if (!buckets[key]) buckets[key] = { key: key, inbound: 0, outbound: 0, returnRecv: 0, damage: 0, returnedAujan: 0, trips: 0 };
          var b = buckets[key];
          b.inbound += +row.inbound_pallets; b.outbound += +row.outbound_pallets; b.returnRecv += +row.return_received_from_delivery;
          b.damage += +row.damage_pallets; b.returnedAujan += +row.pallets_returned_to_aujan; b.trips++;
        });
        return Object.keys(buckets).sort().map(function (k) { return buckets[k]; });
      }
      var monthly = groupBy('month'), weekly = groupBy('week');

      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();

      // ---- Sheet 1: Executive Summary ----
      var es = wb.addWorksheet('Executive Summary');
      es.columns = [{ width: 30 }, { width: 18 }, { width: 40 }];
      styleTitleRow(es, 'AUJAN PALLET INBOUND & OUTBOUND TRACKER — EXECUTIVE SUMMARY', 'C');
      es.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
      es.addRow(['Date Range', (req.query.date_from || 'Earliest') + ' to ' + (req.query.date_to || 'Latest')]);
      es.addRow([]);
      styleHeaderRow(es.addRow(['METRIC', 'VALUE', 'NOTE']));
      function kpiRow(label, val, note, flagRed) {
        var row = es.addRow([label, val, note || '']);
        if (flagRed) { row.getCell(2).font = { bold: true, color: { argb: REDTEXT } }; row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } }; row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } }; }
        else { row.getCell(2).font = { bold: true }; }
      }
      kpiRow('Inbound Pallets (from Aujan)', totals.inbound);
      kpiRow('Outbound Pallets', totals.outbound);
      kpiRow('Return Received from Delivery', totals.returnRecv);
      kpiRow('Damage Pallets', totals.damage, totals.damage > 0 ? 'Damage recorded — review action taken' : '', totals.damage > 0);
      kpiRow('Pallets Returned to Aujan', totals.returnedAujan);
      kpiRow('Customer Retained', customerRetained, customerRetained < 0 ? 'Negative = more sent out than returned from delivery' : '', customerRetained < 0);
      kpiRow('Current Warehouse Balance', currentBalance, 'Running total: Inbound − Outbound + Return Recv. − Damage − Returned to Aujan');
      kpiRow('Total Trips / Entries', rows.length);
      es.addRow([]);
      styleHeaderRow(es.addRow(['LEGEND', '', '']));
      var legendRed = es.addRow(['Red highlight', '', 'Damage recorded, or Customer Retained is negative']);
      legendRed.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } };
      var legendGreen = es.addRow(['Green highlight', '', 'Pallets returned to Aujan on that trip']);
      legendGreen.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREENFLAG } };

      // ---- Sheet 2: Monthly Summary (pivot-style auto-summary) ----
      var ms = wb.addWorksheet('Monthly Summary');
      ms.columns = [{ width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 10 }];
      styleTitleRow(ms, 'MONTHLY BREAKDOWN — select this table in Excel to build a pivot chart', 'H');
      ms.addRow([]);
      styleHeaderRow(ms.addRow(['Month', 'Inbound', 'Outbound', 'Return Recv.', 'Damage', 'Returned to Aujan', 'Net Movement', 'Trips']));
      var mTotals = { inbound: 0, outbound: 0, returnRecv: 0, damage: 0, returnedAujan: 0, trips: 0 };
      monthly.forEach(function (m) {
        var net = m.inbound - m.outbound + m.returnRecv - m.damage - m.returnedAujan;
        var row = ms.addRow([m.key, m.inbound, m.outbound, m.returnRecv, m.damage, m.returnedAujan, net, m.trips]);
        if (m.damage > 0) row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } };
        mTotals.inbound += m.inbound; mTotals.outbound += m.outbound; mTotals.returnRecv += m.returnRecv; mTotals.damage += m.damage; mTotals.returnedAujan += m.returnedAujan; mTotals.trips += m.trips;
      });
      styleTotalRow(ms.addRow(['TOTAL', mTotals.inbound, mTotals.outbound, mTotals.returnRecv, mTotals.damage, mTotals.returnedAujan, mTotals.inbound - mTotals.outbound + mTotals.returnRecv - mTotals.damage - mTotals.returnedAujan, mTotals.trips]));

      // ---- Sheet 3: Weekly Summary ----
      var wsSheet = wb.addWorksheet('Weekly Summary');
      wsSheet.columns = [{ width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 10 }];
      styleTitleRow(wsSheet, 'WEEKLY BREAKDOWN (week starting date)', 'H');
      styleHeaderRow(wsSheet.addRow(['Week Of', 'Inbound', 'Outbound', 'Return Recv.', 'Damage', 'Returned to Aujan', 'Net Movement', 'Trips']));
      weekly.forEach(function (wk) {
        var net = wk.inbound - wk.outbound + wk.returnRecv - wk.damage - wk.returnedAujan;
        var row = wsSheet.addRow([wk.key, wk.inbound, wk.outbound, wk.returnRecv, wk.damage, wk.returnedAujan, net, wk.trips]);
        if (wk.damage > 0) row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } };
      });

      // ---- Sheet 4: Full Ledger (color-coded) ----
      var ws = wb.addWorksheet('Ledger');
      ws.columns = [{ width: 12 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 30 }];
      styleTitleRow(ws, 'AUJAN PALLET INBOUND & OUTBOUND TRACKER — FULL LEDGER', 'K');
      ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
      ws.addRow([]);
      styleHeaderRow(ws.addRow(['Date', 'Truck/Tour/Invoice #', 'Opening', 'Inbound', 'Outbound', 'Return Recv.', 'Damage', 'Returned to Aujan', 'Closing', 'Retained by Customer', 'Remarks / Action Taken']));
      rows.forEach(function (row) {
        var excelRow = ws.addRow([row.entry_date, row.truck_ref, +row.computed_opening_balance, +row.inbound_pallets, +row.outbound_pallets,
        +row.return_received_from_delivery, +row.damage_pallets, +row.pallets_returned_to_aujan, +row.computed_closing_balance,
        +row.pallet_retained_by_customer, row.remarks]);
        if (+row.damage_pallets > 0) {
          excelRow.eachCell(function (cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: REDFLAG } }; });
        } else if (+row.pallets_returned_to_aujan > 0) {
          excelRow.eachCell(function (cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREENFLAG } }; });
        }
      });
      styleTotalRow(ws.addRow(['', 'TOTAL', '', totals.inbound, totals.outbound, totals.returnRecv, totals.damage, totals.returnedAujan, currentBalance, '', '']));

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Aujan_Pallet_Ledger_Analysis_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) {
      console.error('ledger export error:', e.message);
      res.status(500).json({ error: 'Export failed: ' + e.message });
    }
  });

  app.get('/api/pallets/meta', requireAuth, function (req, res) {
    res.json({ statuses: STATUS_LIST, departments: DEPARTMENTS, priorities: PRIORITIES });
  });
};
