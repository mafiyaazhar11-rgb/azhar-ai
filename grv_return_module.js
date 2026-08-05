// ============================================================
// GRV RETURN UPLIFT MODULE — plugs into existing AZHAR-AI server.js
// Self-contained: own table, shares the existing pg pool + auth.
// Purpose: kill the manual header-typing bottleneck (Customer / Shipping
// Location / Salesman) that CS was doing by hand from emailed GRV Uplift
// photos. Merchandisers submit via one shared login + dropdowns (no free
// text, no spelling risk). CS reviews + downloads Excel in the exact
// Alphamed order-form template, then uploads it through the existing
// Excel-upload-to-Oracle flow with Order Type tagged as Return.
//
// Mount with:
//   require('./grv_return_module')(app, pool, requireAuth, requireRole, auditLog);
// ============================================================

const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');

module.exports = function (app, pool, requireAuth, requireRole, auditLog) {

  // ── Init table ──
  async function initGrvDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS grv_return_requests (
        id SERIAL PRIMARY KEY,
        request_ref TEXT UNIQUE NOT NULL,
        account_number TEXT,
        customer_name TEXT,
        site_use_id TEXT,
        location TEXT,
        salesperson_name TEXT NOT NULL,
        warehouse TEXT DEFAULT 'DCF',
        order_type TEXT,
        po_number TEXT,
        requester_email TEXT NOT NULL,
        status TEXT DEFAULT 'Entered',
        submitted_by TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_requests_created ON grv_return_requests(created_at)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_requests_status ON grv_return_requests(status)`);
      console.log('GRV Return module: table ready');
    } catch (e) {
      console.error('GRV Return initDB error:', e.message);
    }
  }
  initGrvDB();

  // ── Helpers ──
  function requireGrvRole(...roles) {
    return function (req, res, next) {
      if (!req.user) return res.status(401).json({ error: 'Not authenticated' });
      if (roles.includes(req.user.role) || req.user.role === 'superadmin' || req.user.role === 'subadmin') return next();
      return res.status(403).json({ error: 'Access denied' });
    };
  }

  async function nextGrvRef() {
    const d = new Date();
    const ymd = d.getFullYear().toString() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    const r = await pool.query(`SELECT COUNT(*)::int AS c FROM grv_return_requests WHERE request_ref LIKE $1`, [`GRV-${ymd}-%`]);
    const seq = (r.rows[0].c || 0) + 1;
    return `GRV-${ymd}-${String(seq).padStart(3, '0')}`;
  }

  // Builds one workbook matching the exact Alphamed order-form template,
  // header fully filled, item grid left blank (15 rows) since this module
  // only ever captures the header — matches the screenshot layout exactly.
  function buildGrvWorkbook(reqRow) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.columns = [
      { width: 20 }, { width: 22 }, { width: 20 }, { width: 34 }, { width: 14 },
      { width: 18 }, { width: 12 }, { width: 10 }, { width: 10 }, { width: 10 }
    ];

    const FILL_ROW1    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0DFE6' } };
    const FILL_GENERAL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA1BFCD' } };
    const FILL_LABEL   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4EA72E' } };
    const FILL_YELLOW  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
    const FILL_ORANGE  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2AA84' } };
    const FILL_SUBCAT  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCFECF7' } };
    const FILL_WHITE   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    for (let r = 1; r <= 6; r++) {
      for (let c = 1; c <= 10; c++) {
        ws.getCell(r, c).fill = r <= 2 ? FILL_ROW1 : FILL_GENERAL;
        ws.getCell(r, c).font = { bold: true };
      }
    }

    ws.mergeCells('A1:J1');
    ws.mergeCells('A2:J2');
    ws.getCell('A1').value = 'ALPHAMED';
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
    ws.getCell('A1').alignment = { horizontal: 'center', wrapText: true };
    ws.getCell('A2').value = 'ORDER FORM';
    ws.getCell('A2').font = { bold: true };
    ws.getCell('A2').fill = FILL_ROW1;
    ws.getCell('A2').alignment = { horizontal: 'center', wrapText: true };

    ws.getCell('A3').value = 'Customer name :';
    ws.getCell('B3').value = 'Customer Nmae : '; ws.getCell('B3').fill = FILL_LABEL;
    ws.getCell('C3').value = reqRow.customer_name || '';
    ws.getCell('F3').value = 'PO:';
    ws.getCell('G3').value = reqRow.po_number || '';

    ws.getCell('A4').value = 'Customer Number  :';
    ws.getCell('B4').value = 'Cutomer code: '; ws.getCell('B4').fill = FILL_LABEL;
    const accountNumVal = reqRow.account_number && !isNaN(reqRow.account_number) ? Number(reqRow.account_number) : (reqRow.account_number || '');
    ws.getCell('C4').value = accountNumVal; ws.getCell('C4').fill = FILL_WHITE; ws.getCell('C4').font = { bold: false };
    ws.getCell('F4').value = 'Sales person name : ' + (reqRow.salesperson_name || '');

    ws.getCell('B5').value = 'W/H :'; ws.getCell('B5').fill = FILL_YELLOW;
    ws.getCell('C5').value = reqRow.warehouse || 'DCF';
    ws.getCell('F5').value = 'Order type';
    ws.getCell('G5').value = reqRow.order_type || '';

    ws.getCell('B6').value = 'Location Site ID :'; ws.getCell('B6').fill = FILL_LABEL;
    const siteIdVal = reqRow.site_use_id && !isNaN(reqRow.site_use_id) ? Number(reqRow.site_use_id) : (reqRow.site_use_id || '');
    ws.getCell('C6').value = siteIdVal;
    ws.getCell('D6').value = 'Drp Down List';
    ws.getCell('F6').value = 'DATE: ' + new Date(reqRow.created_at).toLocaleDateString('en-GB');

    const headerRow = 7;
    const headers = ['Outer Barcode', 'EA Barcode', 'AKI Code', 'Item Description', 'SUB CATEGORY', 'Price per pc/outer', 'Order in PC', 'FOC', 'UOM', 'TOTAL'];
    headers.forEach((h, i) => {
      const cell = ws.getCell(headerRow, 1 + i);
      cell.value = h;
      cell.font = { bold: true };
      if (h === 'SUB CATEGORY') cell.fill = FILL_SUBCAT;
      else if (h === 'Order in PC' || h === 'FOC') cell.fill = FILL_LABEL;
      else cell.fill = FILL_ORANGE;
    });

    // Leave 15 blank line rows — this module only captures the header.
    for (let i = 0; i < 15; i++) {
      const r = headerRow + 1 + i;
      for (let c = 1; c <= 10; c++) ws.getCell(r, c).value = null;
    }

    return wb;
  }

  // ============================================================
  // MERCHANDISER-FACING (shared login, role = 'grv_sales')
  // ============================================================

  // Salesman dropdown list — sourced from the same Customer Master already
  // uploaded for HoReCa (horeca_customer_sites). Open to any authenticated
  // role since the merchandiser must see the full list regardless of customer.
  app.get('/api/grv/salespersons', requireAuth, async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT DISTINCT salesperson_name FROM horeca_customer_sites
         WHERE salesperson_name IS NOT NULL AND salesperson_name != '' ORDER BY salesperson_name`
      );
      res.json({ salespersons: r.rows.map(row => row.salesperson_name) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Submit a return request — header only. No restriction on which
  // customer/door can be picked; salesman must be selected from dropdown.
  app.post('/api/grv/requests', requireAuth, requireGrvRole('grv_sales'), async function (req, res) {
    try {
      const { account_number, customer_name, site_use_id, location, salesperson_name, warehouse, requester_email } = req.body;
      if (!customer_name) return res.status(400).json({ error: 'Please select a customer' });
      if (!salesperson_name) return res.status(400).json({ error: 'Please select the salesman name' });
      if (!requester_email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(requester_email)) {
        return res.status(400).json({ error: 'Please enter a valid email address' });
      }

      const ref = await nextGrvRef();
      const r = await pool.query(
        `INSERT INTO grv_return_requests
           (request_ref, account_number, customer_name, site_use_id, location, salesperson_name, warehouse, requester_email, submitted_by)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING id, request_ref`,
        [ref, account_number || null, customer_name, site_use_id || null, location || null, salesperson_name,
         warehouse || 'DCF', requester_email.trim().toLowerCase(), req.user.username]
      );
      await auditLog(req.user.uid, req.user.username, 'GRV_REQUEST_CREATE', `created ${ref} for ${customer_name}`, '');
      res.json({ success: true, request_ref: r.rows[0].request_ref });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // CS REP-FACING (role = 'grv_cs', or superadmin/subadmin)
  // ============================================================

  // List requests, optionally filtered by date (defaults to today) or 'all'.
  app.get('/api/grv/requests', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const dateFilter = req.query.date; // 'YYYY-MM-DD' or 'all'
      let rows;
      if (dateFilter === 'all') {
        const r = await pool.query(`SELECT * FROM grv_return_requests ORDER BY created_at DESC`);
        rows = r.rows;
      } else {
        const day = dateFilter || new Date().toISOString().slice(0, 10);
        const r = await pool.query(
          `SELECT * FROM grv_return_requests WHERE created_at::date = $1::date ORDER BY created_at DESC`,
          [day]
        );
        rows = r.rows;
      }
      const todayRes = await pool.query(`SELECT COUNT(*)::int AS c FROM grv_return_requests WHERE created_at::date = CURRENT_DATE`);
      res.json({ requests: rows, today_count: todayRes.rows[0].c });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Edit order type / PO / warehouse / status before export.
  app.put('/api/grv/requests/:id', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const { order_type, po_number, warehouse, status } = req.body;
      await pool.query(
        `UPDATE grv_return_requests SET
           order_type = COALESCE($1, order_type),
           po_number  = COALESCE($2, po_number),
           warehouse  = COALESCE($3, warehouse),
           status     = COALESCE($4, status),
           updated_at = NOW()
         WHERE id = $5`,
        [order_type ?? null, po_number ?? null, warehouse ?? null, status ?? null, req.params.id]
      );
      await auditLog(req.user.uid, req.user.username, 'GRV_REQUEST_UPDATE', `updated request #${req.params.id}`, '');
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Download one request as the exact Alphamed template.
  app.get('/api/grv/requests/:id/export-excel', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const r = await pool.query(`SELECT * FROM grv_return_requests WHERE id=$1`, [req.params.id]);
      const reqRow = r.rows[0];
      if (!reqRow) return res.status(404).json({ error: 'Request not found' });

      const wb = buildGrvWorkbook(reqRow);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${reqRow.request_ref}.xlsx"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // Download all requests for a given day (default today) as a ZIP of
  // individual template files — the CS end-of-day batch download.
  app.get('/api/grv/requests/export-all', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const dateFilter = req.query.date;
      let rows;
      if (dateFilter === 'all') {
        const r = await pool.query(`SELECT * FROM grv_return_requests ORDER BY created_at DESC`);
        rows = r.rows;
      } else {
        const day = dateFilter || new Date().toISOString().slice(0, 10);
        const r = await pool.query(`SELECT * FROM grv_return_requests WHERE created_at::date = $1::date ORDER BY created_at DESC`, [day]);
        rows = r.rows;
      }
      if (!rows.length) return res.status(404).json({ error: 'No requests found for that date' });

      const zip = new AdmZip();
      for (const reqRow of rows) {
        const wb = buildGrvWorkbook(reqRow);
        const buf = await wb.xlsx.writeBuffer();
        zip.addFile(`${reqRow.request_ref}.xlsx`, buf);
      }
      const zipBuf = zip.toBuffer();
      const label = dateFilter === 'all' ? 'ALL' : (dateFilter || new Date().toISOString().slice(0, 10));
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', `attachment; filename="GRV_Returns_${label}.zip"`);
      res.send(zipBuf);
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  console.log('GRV Return module: routes mounted');
};
