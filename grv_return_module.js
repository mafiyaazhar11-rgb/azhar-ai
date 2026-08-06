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
// GRV has its OWN customer/site/salesman master — separate business unit
// from HoReCa, deliberately not sharing horeca_customer_sites.
//
// Mount with:
//   require('./grv_return_module')(app, pool, requireAuth, requireRole, upload, auditLog);
// ============================================================

const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog) {

  // ── Init tables ──
  async function initGrvDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS grv_customer_sites (
        id SERIAL PRIMARY KEY,
        account_number TEXT NOT NULL,
        customer_name TEXT NOT NULL,
        customer_category TEXT,
        salesperson_code TEXT,
        salesperson_name TEXT,
        location TEXT,
        address_detail TEXT,
        site_use_id TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(account_number, site_use_id)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_sites_account ON grv_customer_sites(account_number)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_sites_salesperson ON grv_customer_sites(salesperson_name)`);

      // Standalone Salesperson Master — deliberately NOT tied to any customer.
      // A merchandiser can pick any salesman for any customer; the list here
      // is just the full roster to choose from.
      await pool.query(`CREATE TABLE IF NOT EXISTS grv_salespersons (
        id SERIAL PRIMARY KEY,
        salesperson_code TEXT,
        salesperson_name TEXT UNIQUE NOT NULL,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);

      await pool.query(`CREATE TABLE IF NOT EXISTS grv_return_requests (
        id SERIAL PRIMARY KEY,
        request_ref TEXT UNIQUE NOT NULL,
        account_number TEXT,
        customer_name TEXT,
        site_use_id TEXT,
        location TEXT,
        salesperson_name TEXT NOT NULL,
        warehouse TEXT,
        order_type TEXT,
        po_number TEXT,
        requester_email TEXT NOT NULL,
        status TEXT DEFAULT 'Entered',
        submitted_by TEXT,
        assigned_to INT REFERENCES users(id),
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`ALTER TABLE grv_return_requests ADD COLUMN IF NOT EXISTS assigned_to INT REFERENCES users(id)`);
      await pool.query(`ALTER TABLE grv_return_requests ADD COLUMN IF NOT EXISTS paper_return_number TEXT`);
      await pool.query(`ALTER TABLE grv_return_requests ALTER COLUMN warehouse DROP DEFAULT`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_requests_created ON grv_return_requests(created_at)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_requests_status ON grv_return_requests(status)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS grv_items (
        id SERIAL PRIMARY KEY,
        item_code TEXT UNIQUE NOT NULL,
        description TEXT,
        brand TEXT,
        category TEXT,
        uom TEXT,
        price_per_unit NUMERIC,
        ea_barcode TEXT,
        outer_barcode TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`ALTER TABLE grv_items ADD COLUMN IF NOT EXISTS ea_barcode TEXT`);
      await pool.query(`ALTER TABLE grv_items ADD COLUMN IF NOT EXISTS outer_barcode TEXT`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_items_desc ON grv_items USING gin (to_tsvector('simple', coalesce(description,'') || ' ' || coalesce(item_code,'')))`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_items_ea_barcode ON grv_items(ea_barcode)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_items_outer_barcode ON grv_items(outer_barcode)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS grv_return_request_lines (
        id SERIAL PRIMARY KEY,
        request_id INT REFERENCES grv_return_requests(id) ON DELETE CASCADE,
        item_code TEXT,
        description TEXT,
        sub_category TEXT,
        uom TEXT,
        price_per_unit NUMERIC,
        qty NUMERIC,
        foc_qty NUMERIC DEFAULT 0,
        reason_code INT,
        reason_label TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_grv_lines_request ON grv_return_request_lines(request_id)`);

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

  // Fixed return-reason key, matches the printed key on the physical GRV Uplift Form.
  const GRV_REASON_CODES = {
    1: 'Box Damage', 2: 'Wrong Delivery', 3: 'Promotion Return', 4: 'Slow Moving',
    5: 'Expired', 6: 'Wrong Price', 7: 'Other', 8: 'Wrong Barcode'
  };

  // Builds one workbook matching the exact Alphamed order-form template,
  // header fully filled, item grid populated from the request's line items.
  function buildGrvWorkbook(reqRow, lines) {
    lines = lines || [];
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
    ws.getCell('C5').value = reqRow.warehouse || '';
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

    // Populate item grid from the request's actual line items.
    lines.forEach((l, idx) => {
      const r = headerRow + 1 + idx;
      const codeCell = ws.getCell(r, 3); codeCell.value = l.item_code || ''; codeCell.font = { bold: false };  // C = AKI Code
      ws.getCell(r, 4).value = l.description || '';                                                             // D = Item Description
      ws.getCell(r, 5).value = l.sub_category || '';                                                            // E = SUB CATEGORY
      if (l.price_per_unit != null) ws.getCell(r, 6).value = Number(l.price_per_unit);                          // F = Price per pc/outer
      const qtyCell = ws.getCell(r, 7); qtyCell.value = l.qty != null ? Number(l.qty) : null; qtyCell.font = { bold: true }; qtyCell.alignment = { horizontal: 'center' }; // G = Order in PC
      if (l.foc_qty != null) ws.getCell(r, 8).value = Number(l.foc_qty);                                        // H = FOC
      ws.getCell(r, 9).value = l.uom || '';                                                                     // I = UOM
      const totalCell = ws.getCell(r, 10); totalCell.value = '-'; totalCell.font = { bold: true };              // J = TOTAL (Oracle-calculated, placeholder)
    });

    return wb;
  }

  // ============================================================
  // ADMIN: GRV'S OWN MASTER DATA UPLOAD (add/update only, never delete)
  // Separate business unit from HoReCa — same Excel column layout for
  // convenience (Account Number, Site Use Id, Customer Name, Customer
  // Category, Salesmen Number, New Salespersons, Location, Address2/3, City)
  // but stored in its own grv_customer_sites table.
  // ============================================================

  app.post('/api/grv/master/customers/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames.find(n => /customar master|customer master/i.test(n)) || wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

      const map = new Map();
      let skipped = 0, rowCounter = 0;
      for (const row of rows) {
        rowCounter++;
        const accountNumber = String(row['Account Number'] || row['Account_Number'] || '').trim();
        const siteUseIdRaw = String(row['Site Use Id'] || row['Site Use ID'] || row['Ship to Site Use ID'] || row['Ship to Site Use Id'] || '').trim();
        const customerName = String(row['Customer Name'] || row['Customer_Name'] || '').trim();
        if (!accountNumber || !customerName) { skipped++; continue; }

        const category = row['Customer Category'] || null;
        const spCode = row['Salesmen Number'] != null ? String(row['Salesmen Number']).trim() : null;
        const spName = row['New Salespersons'] != null ? String(row['New Salespersons']).trim() : null;
        const location = row['Location'] || row['Ship to Location'] || null;

        const addressParts = [row['Address2'], row['Address3'], row['City']]
          .map(v => (v == null ? '' : String(v).trim()))
          .filter(v => v.length > 0);
        const addressDetail = addressParts.length ? addressParts.join(', ') : null;

        // If Site Use Id is blank, derive a stable key from the location text
        // instead of leaving it blank — otherwise every blank-Site-Use-Id row
        // for the same customer collides on the same key and silently
        // overwrites the previous location instead of adding a new one.
        const siteUseId = siteUseIdRaw || (location ? 'LOC_' + String(location).trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_') : `ROW_${rowCounter}`);

        map.set(accountNumber + '||' + siteUseId, [accountNumber, customerName, category, spCode, spName, location, addressDetail, siteUseId]);
      }
      const records = Array.from(map.values());

      let inserted = 0, updated = 0;
      const BATCH = 500;
      for (let i = 0; i < records.length; i += BATCH) {
        const chunk = records.slice(i, i + BATCH);
        const valuesSql = [];
        const params = [];
        chunk.forEach((rec, idx) => {
          const b = idx * 8;
          valuesSql.push(`($${b+1},$${b+2},$${b+3},$${b+4},$${b+5},$${b+6},$${b+7},$${b+8})`);
          params.push(...rec);
        });
        const r = await pool.query(
          `INSERT INTO grv_customer_sites (account_number, customer_name, customer_category, salesperson_code, salesperson_name, location, address_detail, site_use_id)
           VALUES ${valuesSql.join(',')}
           ON CONFLICT (account_number, site_use_id) DO UPDATE SET
             customer_name=EXCLUDED.customer_name, customer_category=EXCLUDED.customer_category,
             salesperson_code=EXCLUDED.salesperson_code, salesperson_name=EXCLUDED.salesperson_name,
             location=EXCLUDED.location, address_detail=EXCLUDED.address_detail, updated_at=NOW()
           RETURNING (xmax = 0) AS was_insert`,
          params
        );
        for (const row of r.rows) { if (row.was_insert) inserted++; else updated++; }
      }

      await auditLog(req.user.uid, req.user.username, 'GRV_CUSTOMER_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Export currently-stored GRV Customer Master (for admin to verify what's saved)
  app.get('/api/grv/master/customers/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT account_number AS "Account Number", customer_name AS "Customer Name", customer_category AS "Customer Category",
                salesperson_name AS "New Salespersons", salesperson_code AS "Salesmen Number", location AS "Location", site_use_id AS "Site Use Id"
         FROM grv_customer_sites WHERE active=true ORDER BY customer_name`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Customar Master');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 22 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="GRV_Customer_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // ADMIN: STANDALONE SALESPERSON MASTER (not tied to any customer)
  // ============================================================

  app.post('/api/grv/master/salespersons/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

      const map = new Map();
      let skipped = 0;
      for (const row of rows) {
        const name = String(row['New Salespersons'] || row['Salesperson Name'] || row['Salesman Name'] || row['Name'] || '').trim();
        if (!name) { skipped++; continue; }
        const code = row['Salesmen Number'] != null ? String(row['Salesmen Number']).trim() : (row['Code'] != null ? String(row['Code']).trim() : null);
        map.set(name, [code, name]);
      }
      const records = Array.from(map.values());

      let inserted = 0, updated = 0;
      const BATCH = 500;
      for (let i = 0; i < records.length; i += BATCH) {
        const chunk = records.slice(i, i + BATCH);
        const valuesSql = [];
        const params = [];
        chunk.forEach((rec, idx) => {
          const b = idx * 2;
          valuesSql.push(`($${b+1},$${b+2})`);
          params.push(...rec);
        });
        const r = await pool.query(
          `INSERT INTO grv_salespersons (salesperson_code, salesperson_name)
           VALUES ${valuesSql.join(',')}
           ON CONFLICT (salesperson_name) DO UPDATE SET salesperson_code=EXCLUDED.salesperson_code, updated_at=NOW()
           RETURNING (xmax = 0) AS was_insert`,
          params
        );
        for (const row of r.rows) { if (row.was_insert) inserted++; else updated++; }
      }

      await auditLog(req.user.uid, req.user.username, 'GRV_SALESPERSON_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.get('/api/grv/master/salespersons/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT salesperson_code AS "Salesmen Number", salesperson_name AS "New Salespersons"
         FROM grv_salespersons WHERE active=true ORDER BY salesperson_name`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Salespersons');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 26 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="GRV_Salesperson_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // Upload Item Master (flexible columns: Item_Code/Item Code/AKI Code, Description, Brand, Category, UOM, Price,
  // and barcode as either a single "Barcode" column or separate "EA Barcode"/"Outer Barcode" columns).
  // Batched upsert — this file can be tens of thousands of rows, so we never
  // do a per-row SELECT+INSERT/UPDATE round trip (that would time out on a
  // 50k-row file). Rows are deduped by item_code (last occurrence wins,
  // same as before) and sent to Postgres in chunks via ON CONFLICT.
  app.post('/api/grv/master/items/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
      const cleanBarcode = (v) => v == null ? null : String(v).trim().replace(/^'+/, '') || null;

      const map = new Map();
      let skipped = 0;
      for (const row of rows) {
        const itemCode = String(row['Item_Code'] || row['Item Code'] || row['AKI Code'] || '').trim();
        if (!itemCode) { skipped++; continue; }
        const description = row['Description'] || row['Item Description'] || null;
        const brand = row['Brand'] || null;
        const category = row['Category'] || row['SUB CATEGORY'] || null;
        const uom = row['UOM'] || null;
        const priceRaw = row['Price'] || row['Price per pc/outer'] || null;
        const price = priceRaw != null && !isNaN(priceRaw) ? Number(priceRaw) : null;
        // Prefer explicit EA/Outer columns if present; otherwise fall back to a
        // single generic "Barcode" column (used as the EA/unit barcode).
        const eaBarcode = cleanBarcode(row['EA Barcode'] != null ? row['EA Barcode'] : row['Barcode']);
        const outerBarcode = cleanBarcode(row['Outer Barcode']);
        map.set(itemCode, [itemCode, description, brand, category, uom, price, eaBarcode, outerBarcode]);
      }
      const records = Array.from(map.values());

      let inserted = 0, updated = 0;
      const BATCH = 500;
      for (let i = 0; i < records.length; i += BATCH) {
        const chunk = records.slice(i, i + BATCH);
        const valuesSql = [];
        const params = [];
        chunk.forEach((rec, idx) => {
          const b = idx * 8;
          valuesSql.push(`($${b+1},$${b+2},$${b+3},$${b+4},$${b+5},$${b+6},$${b+7},$${b+8})`);
          params.push(...rec);
        });
        const r = await pool.query(
          `INSERT INTO grv_items (item_code, description, brand, category, uom, price_per_unit, ea_barcode, outer_barcode)
           VALUES ${valuesSql.join(',')}
           ON CONFLICT (item_code) DO UPDATE SET
             description=EXCLUDED.description, brand=EXCLUDED.brand, category=EXCLUDED.category,
             uom=EXCLUDED.uom, price_per_unit=EXCLUDED.price_per_unit, ea_barcode=EXCLUDED.ea_barcode,
             outer_barcode=EXCLUDED.outer_barcode, updated_at=NOW()
           RETURNING (xmax = 0) AS was_insert`,
          params
        );
        for (const row of r.rows) { if (row.was_insert) inserted++; else updated++; }
      }

      await auditLog(req.user.uid, req.user.username, 'GRV_ITEM_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Export currently-stored GRV Item Master
  app.get('/api/grv/master/items/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT item_code AS "Item Code", description AS "Description", brand AS "Brand", category AS "Category",
                uom AS "UOM", price_per_unit AS "Price", ea_barcode AS "EA Barcode", outer_barcode AS "Outer Barcode"
         FROM grv_items WHERE active=true ORDER BY item_code`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Items');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 22 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="GRV_Item_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // Item search for the merchandiser's line-item picker (code, description, or barcode).
  // Item search — used by both the combined search and the two dedicated
  // Item Code / Description fields on the merchandiser form.
  // field=code -> item_code only, field=desc -> description only, omitted -> all.
  app.get('/api/grv/items', requireAuth, async function (req, res) {
    try {
      const q = (req.query.q || '').trim();
      const field = req.query.field;
      if (!q) return res.json({ items: [] });
      let whereClause;
      if (field === 'code') whereClause = `item_code ILIKE $1`;
      else if (field === 'desc') whereClause = `description ILIKE $1`;
      else whereClause = `(item_code ILIKE $1 OR description ILIKE $1 OR ea_barcode ILIKE $1 OR outer_barcode ILIKE $1)`;
      const r = await pool.query(
        `SELECT item_code, description, brand, category, uom, price_per_unit, ea_barcode, outer_barcode FROM grv_items
         WHERE active=true AND ${whereClause}
         ORDER BY item_code LIMIT 20`,
        [`%${q}%`]
      );
      res.json({ items: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Exact barcode lookup — used by the camera scanner. Checks both EA and
  // Outer barcode, since a merchandiser might scan either.
  app.get('/api/grv/items/barcode', requireAuth, async function (req, res) {
    try {
      const code = (req.query.code || '').trim();
      if (!code) return res.status(400).json({ error: 'No barcode provided' });
      const r = await pool.query(
        `SELECT item_code, description, brand, category, uom, price_per_unit, ea_barcode, outer_barcode FROM grv_items
         WHERE active=true AND (ea_barcode = $1 OR outer_barcode = $1) LIMIT 1`,
        [code]
      );
      if (!r.rows.length) return res.status(404).json({ error: 'No item found for this barcode' });
      res.json({ item: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // MERCHANDISER-FACING (shared login, role = 'grv_sales')
  // ============================================================

  // Customers + sites grouped by account_number — sourced from GRV's own
  // master. No restriction on which customer/door can be picked, per Azhar's
  // requirement ("no limitation all can see any customer").
  app.get('/api/grv/customers', requireAuth, async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT account_number, customer_name, customer_category, location, address_detail, site_use_id
         FROM grv_customer_sites WHERE active=true ORDER BY customer_name`
      );
      const grouped = {};
      for (const row of r.rows) {
        if (!grouped[row.account_number]) {
          grouped[row.account_number] = {
            account_number: row.account_number,
            customer_name: row.customer_name,
            customer_category: row.customer_category,
            sites: []
          };
        }
        grouped[row.account_number].sites.push({ location: row.location, address_detail: row.address_detail, site_use_id: row.site_use_id });
      }
      res.json({ customers: Object.values(grouped) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Salesman dropdown list — sourced from the standalone Salesperson Master,
  // deliberately independent of customer selection (any salesman, any customer).
  app.get('/api/grv/salespersons', requireAuth, async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT salesperson_name FROM grv_salespersons WHERE active=true ORDER BY salesperson_name`
      );
      res.json({ salespersons: r.rows.map(row => row.salesperson_name) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Submit a return request — header + line items. No restriction on which
  // customer/door can be picked; salesman must be selected from dropdown.
  app.post('/api/grv/requests', requireAuth, requireGrvRole('grv_sales'), async function (req, res) {
    try {
      const { account_number, customer_name, site_use_id, location, salesperson_name, warehouse, requester_email, lines, paper_return_number } = req.body;
      if (!customer_name) return res.status(400).json({ error: 'Please select a customer' });
      if (!salesperson_name) return res.status(400).json({ error: 'Please select the salesman name' });
      if (!paper_return_number || !paper_return_number.trim()) {
        return res.status(400).json({ error: 'Please enter the Return Number printed on the physical GRV form' });
      }
      if (!requester_email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(requester_email)) {
        return res.status(400).json({ error: 'Please enter a valid email address' });
      }
      if (!Array.isArray(lines) || !lines.length) {
        return res.status(400).json({ error: 'Please add at least one item to the return request' });
      }
      for (const l of lines) {
        if (!l.item_code) return res.status(400).json({ error: 'Every line must have an item selected' });
        if (l.qty == null || isNaN(l.qty) || Number(l.qty) <= 0) return res.status(400).json({ error: `Enter a valid return quantity for ${l.item_code}` });
      }

      const ref = await nextGrvRef();
      const r = await pool.query(
        `INSERT INTO grv_return_requests
           (request_ref, account_number, customer_name, site_use_id, location, salesperson_name, warehouse, requester_email, submitted_by, paper_return_number)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) RETURNING id, request_ref`,
        [ref, account_number || null, customer_name, site_use_id || null, location || null, salesperson_name,
         warehouse || null, requester_email.trim().toLowerCase(), req.user.username, paper_return_number.trim()]
      );
      const requestId = r.rows[0].id;

      for (const l of lines) {
        const reasonCode = l.reason_code != null ? Number(l.reason_code) : null;
        const reasonLabel = reasonCode && GRV_REASON_CODES[reasonCode] ? GRV_REASON_CODES[reasonCode] : (l.reason_label || null);
        await pool.query(
          `INSERT INTO grv_return_request_lines (request_id, item_code, description, sub_category, uom, price_per_unit, qty, foc_qty, reason_code, reason_label)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`,
          [requestId, l.item_code, l.description || null, l.category || null, l.uom || null,
           l.price_per_unit != null ? Number(l.price_per_unit) : null, Number(l.qty), l.foc_qty != null ? Number(l.foc_qty) : 0,
           reasonCode, reasonLabel]
        );
      }

      await auditLog(req.user.uid, req.user.username, 'GRV_REQUEST_CREATE', `created ${ref} for ${customer_name} (${lines.length} lines)`, '');
      res.json({ success: true, request_ref: r.rows[0].request_ref });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // CS REP-FACING (role = 'grv_cs', or superadmin/subadmin)
  // ============================================================

  // Users who can be assigned a request — anyone with the grv_cs role,
  // plus superadmin/subadmin who also have access to this view.
  app.get('/api/grv/cs-users', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT id, username, full_name FROM users
         WHERE role IN ('grv_cs','superadmin','subadmin') AND active=true ORDER BY full_name`
      );
      res.json({ users: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // List requests, optionally filtered by date (defaults to today) or 'all'.
  // Each request includes its line items nested under "lines", plus who
  // it's assigned to (assigned_to / assigned_to_name).
  app.get('/api/grv/requests', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const dateFilter = req.query.date; // 'YYYY-MM-DD' or 'all'
      let rows;
      if (dateFilter === 'all') {
        const r = await pool.query(
          `SELECT gr.*, u.full_name AS assigned_to_name FROM grv_return_requests gr
           LEFT JOIN users u ON u.id = gr.assigned_to ORDER BY gr.created_at DESC`
        );
        rows = r.rows;
      } else {
        const day = dateFilter || new Date().toISOString().slice(0, 10);
        const r = await pool.query(
          `SELECT gr.*, u.full_name AS assigned_to_name FROM grv_return_requests gr
           LEFT JOIN users u ON u.id = gr.assigned_to
           WHERE gr.created_at::date = $1::date ORDER BY gr.created_at DESC`,
          [day]
        );
        rows = r.rows;
      }
      if (rows.length) {
        const ids = rows.map(r => r.id);
        const linesRes = await pool.query(`SELECT * FROM grv_return_request_lines WHERE request_id = ANY($1::int[]) ORDER BY id`, [ids]);
        const linesByReq = {};
        for (const l of linesRes.rows) {
          if (!linesByReq[l.request_id]) linesByReq[l.request_id] = [];
          linesByReq[l.request_id].push(l);
        }
        rows.forEach(r => { r.lines = linesByReq[r.id] || []; });
      }
      const todayRes = await pool.query(`SELECT COUNT(*)::int AS c FROM grv_return_requests WHERE created_at::date = CURRENT_DATE`);
      res.json({ requests: rows, today_count: todayRes.rows[0].c });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Edit order type / PO / warehouse / status / assigned CS person before export.
  app.put('/api/grv/requests/:id', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const { order_type, po_number, warehouse, status, assigned_to, paper_return_number } = req.body;
      const assignedToProvided = Object.prototype.hasOwnProperty.call(req.body, 'assigned_to');
      const assignedToVal = assigned_to === '' || assigned_to == null ? null : Number(assigned_to);

      if (assignedToProvided) {
        await pool.query(
          `UPDATE grv_return_requests SET
             order_type  = COALESCE($1, order_type),
             po_number   = COALESCE($2, po_number),
             warehouse   = COALESCE($3, warehouse),
             status      = COALESCE($4, status),
             assigned_to = $5,
             paper_return_number = COALESCE($6, paper_return_number),
             updated_at  = NOW()
           WHERE id = $7`,
          [order_type ?? null, po_number ?? null, warehouse ?? null, status ?? null, assignedToVal, paper_return_number ?? null, req.params.id]
        );
      } else {
        await pool.query(
          `UPDATE grv_return_requests SET
             order_type = COALESCE($1, order_type),
             po_number  = COALESCE($2, po_number),
             warehouse  = COALESCE($3, warehouse),
             status     = COALESCE($4, status),
             paper_return_number = COALESCE($5, paper_return_number),
             updated_at = NOW()
           WHERE id = $6`,
          [order_type ?? null, po_number ?? null, warehouse ?? null, status ?? null, paper_return_number ?? null, req.params.id]
        );
      }
      await auditLog(req.user.uid, req.user.username, 'GRV_REQUEST_UPDATE', `updated request #${req.params.id}`, '');
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Download one request as the exact Alphamed template, with real line items.
  app.get('/api/grv/requests/:id/export-excel', requireAuth, requireGrvRole('grv_cs'), async function (req, res) {
    try {
      const r = await pool.query(`SELECT * FROM grv_return_requests WHERE id=$1`, [req.params.id]);
      const reqRow = r.rows[0];
      if (!reqRow) return res.status(404).json({ error: 'Request not found' });
      const linesRes = await pool.query(`SELECT * FROM grv_return_request_lines WHERE request_id=$1 ORDER BY id`, [req.params.id]);

      const wb = buildGrvWorkbook(reqRow, linesRes.rows);
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
        const linesRes = await pool.query(`SELECT * FROM grv_return_request_lines WHERE request_id=$1 ORDER BY id`, [reqRow.id]);
        const wb = buildGrvWorkbook(reqRow, linesRes.rows);
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
