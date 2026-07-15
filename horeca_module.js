// ============================================================
// HoReCa ORDER MODULE — plugs into existing AZHAR-AI server.js
// Self-contained: own tables, shares the existing pg pool + auth.
// Mount with: require('./horeca_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt, crypto);
// ============================================================

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, bcrypt, crypto) {

  const MAX_LINES_PER_ORDER = 10;
  const MAX_ORDERS_PER_DAY = 40;

  // ── Init tables ──
  async function initHorecaDB() {
    try {
      await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS horeca_salesperson_code TEXT`);
      await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS horeca_salesperson_name TEXT`);

      await pool.query(`CREATE TABLE IF NOT EXISTS horeca_customer_sites (
        id SERIAL PRIMARY KEY,
        account_number TEXT NOT NULL,
        customer_name TEXT NOT NULL,
        customer_category TEXT,
        salesperson_code TEXT,
        salesperson_name TEXT,
        location TEXT,
        site_use_id TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(account_number, site_use_id)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_sites_account ON horeca_customer_sites(account_number)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_sites_salesperson ON horeca_customer_sites(salesperson_code)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS horeca_items (
        id SERIAL PRIMARY KEY,
        item_code TEXT UNIQUE NOT NULL,
        description TEXT,
        brand TEXT,
        category TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_items_desc ON horeca_items USING gin (to_tsvector('simple', coalesce(description,'') || ' ' || coalesce(item_code,'')))`);

      await pool.query(`CREATE TABLE IF NOT EXISTS horeca_orders (
        id SERIAL PRIMARY KEY,
        order_ref TEXT UNIQUE NOT NULL,
        salesman_user_id INT REFERENCES users(id),
        salesperson_name TEXT,
        account_number TEXT,
        customer_name TEXT,
        site_use_id TEXT,
        location TEXT,
        po_number TEXT,
        warehouse TEXT DEFAULT 'DCF',
        status TEXT DEFAULT 'pending',
        created_at TIMESTAMPTZ DEFAULT NOW(),
        actioned_at TIMESTAMPTZ,
        actioned_by TEXT
      )`);
      await pool.query(`ALTER TABLE horeca_orders ADD COLUMN IF NOT EXISTS warehouse TEXT DEFAULT 'DCF'`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_orders_salesman ON horeca_orders(salesman_user_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_orders_created ON horeca_orders(created_at)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS horeca_order_lines (
        id SERIAL PRIMARY KEY,
        order_id INT REFERENCES horeca_orders(id) ON DELETE CASCADE,
        item_code TEXT,
        description TEXT,
        qty NUMERIC,
        uom TEXT
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_lines_order ON horeca_order_lines(order_id)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS horeca_notifications (
        id SERIAL PRIMARY KEY,
        user_id INT REFERENCES users(id),
        order_id INT REFERENCES horeca_orders(id),
        message TEXT,
        read BOOLEAN DEFAULT false,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_horeca_notif_user ON horeca_notifications(user_id)`);

      console.log('HoReCa module: tables ready');
    } catch (e) {
      console.error('HoReCa initDB error:', e.message);
    }
  }
  initHorecaDB();

  // ── Helpers ──
  function requireHorecaRole(...roles) {
    return function (req, res, next) {
      if (!req.user) return res.status(401).json({ error: 'Not authenticated' });
      if (roles.includes(req.user.role) || req.user.role === 'superadmin') return next();
      return res.status(403).json({ error: 'Access denied' });
    };
  }

  async function nextOrderRef() {
    const d = new Date();
    const ymd = d.getFullYear().toString() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    const r = await pool.query(`SELECT COUNT(*)::int AS c FROM horeca_orders WHERE order_ref LIKE $1`, [`HO-${ymd}-%`]);
    const seq = (r.rows[0].c || 0) + 1;
    return `HO-${ymd}-${String(seq).padStart(3, '0')}`;
  }

  // ============================================================
  // ADMIN: MASTER DATA UPLOAD (add/update only, never delete)
  // ============================================================

  // Upload Customer Master (expects the "Customar Master" sheet layout)
  app.post('/api/horeca/master/customers/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames.find(n => /customar master|customer master/i.test(n)) || wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

      let inserted = 0, updated = 0, skipped = 0;
      for (const row of rows) {
        const accountNumber = String(row['Account Number'] || '').trim();
        const siteUseId = String(row['Site Use Id'] || row['Site Use ID'] || '').trim();
        const customerName = String(row['Customer Name'] || '').trim();
        if (!accountNumber || !customerName) { skipped++; continue; }

        const category = row['Customer Category'] || null;
        const spCode = row['Salesmen Number'] != null ? String(row['Salesmen Number']).trim() : null;
        const spName = row['New Salespersons'] != null ? String(row['New Salespersons']).trim() : null;
        const location = row['Location'] || null;

        const existing = await pool.query(
          `SELECT id FROM horeca_customer_sites WHERE account_number=$1 AND site_use_id=$2`,
          [accountNumber, siteUseId || '']
        );
        if (existing.rows.length) {
          await pool.query(
            `UPDATE horeca_customer_sites SET customer_name=$1, customer_category=$2, salesperson_code=$3, salesperson_name=$4, location=$5, updated_at=NOW()
             WHERE id=$6`,
            [customerName, category, spCode, spName, location, existing.rows[0].id]
          );
          updated++;
        } else {
          await pool.query(
            `INSERT INTO horeca_customer_sites (account_number, customer_name, customer_category, salesperson_code, salesperson_name, location, site_use_id)
             VALUES ($1,$2,$3,$4,$5,$6,$7)`,
            [accountNumber, customerName, category, spCode, spName, location, siteUseId || '']
          );
          inserted++;
        }
      }
      await auditLog(req.user.uid, req.user.username, 'HORECA_CUSTOMER_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Upload Item Master (flexible columns: Item_Code/Item Code, Description, Brand, Category)
  app.post('/api/horeca/master/items/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

      let inserted = 0, updated = 0, skipped = 0;
      for (const row of rows) {
        const itemCode = String(row['Item_Code'] || row['Item Code'] || row['AKI Code'] || '').trim();
        if (!itemCode) { skipped++; continue; }
        const description = row['Description'] || row['Item Description'] || null;
        const brand = row['Brand'] || null;
        const category = row['LAH Category'] || row['Category'] || row['SUB CATEGORY'] || null;

        const existing = await pool.query(`SELECT id FROM horeca_items WHERE item_code=$1`, [itemCode]);
        if (existing.rows.length) {
          await pool.query(
            `UPDATE horeca_items SET description=$1, brand=$2, category=$3, updated_at=NOW() WHERE id=$4`,
            [description, brand, category, existing.rows[0].id]
          );
          updated++;
        } else {
          await pool.query(
            `INSERT INTO horeca_items (item_code, description, brand, category) VALUES ($1,$2,$3,$4)`,
            [itemCode, description, brand, category]
          );
          inserted++;
        }
      }
      await auditLog(req.user.uid, req.user.username, 'HORECA_ITEM_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Export currently-stored Customer Master (for admin to verify what's actually saved)
  app.get('/api/horeca/master/customers/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT account_number AS "Account Number", customer_name AS "Customer Name", customer_category AS "Customer Category",
                salesperson_name AS "New Salespersons", salesperson_code AS "Salesmen Number", location AS "Location", site_use_id AS "Site Use Id"
         FROM horeca_customer_sites WHERE active=true ORDER BY customer_name`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Customar Master');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 22 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="HoReCa_Customer_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // Export currently-stored Item Master
  app.get('/api/horeca/master/items/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT item_code AS "Item Code", description AS "Description", brand AS "Brand", category AS "Category"
         FROM horeca_items WHERE active=true ORDER BY item_code`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Items');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 22 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="HoReCa_Item_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // SALESMAN: customers, items, order submission, own status
  // ============================================================

  // Customers assigned to the logged-in salesman (grouped by account_number so UI can show sites)
  app.get('/api/horeca/customers', requireAuth, async function (req, res) {
    try {
      let rows;
      if (req.user.role === 'salesman') {
        const name = req.user.horeca_salesperson_name;
        if (!name) return res.json({ customers: [] });
        const r = await pool.query(
          `SELECT account_number, customer_name, customer_category, location, site_use_id
           FROM horeca_customer_sites WHERE salesperson_name=$1 AND active=true ORDER BY customer_name`,
          [name]
        );
        rows = r.rows;
      } else {
        const r = await pool.query(
          `SELECT account_number, customer_name, customer_category, location, site_use_id, salesperson_name
           FROM horeca_customer_sites WHERE active=true ORDER BY customer_name`
        );
        rows = r.rows;
      }
      // group by account_number -> sites[]
      const grouped = {};
      for (const row of rows) {
        if (!grouped[row.account_number]) {
          grouped[row.account_number] = {
            account_number: row.account_number,
            customer_name: row.customer_name,
            customer_category: row.customer_category,
            sites: []
          };
        }
        grouped[row.account_number].sites.push({ location: row.location, site_use_id: row.site_use_id });
      }
      res.json({ customers: Object.values(grouped) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Item search (code or description)
  app.get('/api/horeca/items', requireAuth, async function (req, res) {
    try {
      const q = (req.query.q || '').trim();
      let r;
      if (!q) {
        r = await pool.query(`SELECT item_code, description, brand, category FROM horeca_items WHERE active=true ORDER BY item_code LIMIT 50`);
      } else {
        r = await pool.query(
          `SELECT item_code, description, brand, category FROM horeca_items
           WHERE active=true AND (item_code ILIKE $1 OR description ILIKE $1) ORDER BY item_code LIMIT 50`,
          [`%${q}%`]
        );
      }
      res.json({ items: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Submit a new order
  app.post('/api/horeca/orders', requireAuth, requireHorecaRole('salesman'), async function (req, res) {
    const client = await pool.connect();
    try {
      const { account_number, customer_name, site_use_id, location, po_number, warehouse, lines } = req.body;
      if (!account_number || !customer_name) return res.status(400).json({ error: 'Customer is required' });
      if (!Array.isArray(lines) || lines.length === 0) return res.status(400).json({ error: 'At least one order line is required' });
      if (lines.length > MAX_LINES_PER_ORDER) return res.status(400).json({ error: `Maximum ${MAX_LINES_PER_ORDER} lines per order` });
      for (const l of lines) {
        if (!l.item_code || !l.qty || !l.uom) return res.status(400).json({ error: 'Each line needs item code, qty and UOM' });
        if (!['EA', 'CS', 'BNS'].includes(String(l.uom).toUpperCase())) return res.status(400).json({ error: 'UOM must be EA, CS or BNS' });
      }

      // soft daily-volume check (warning only, does not block)
      const today = new Date(); today.setHours(0,0,0,0);
      const countRes = await client.query(`SELECT COUNT(*)::int AS c FROM horeca_orders WHERE created_at >= $1`, [today.toISOString()]);
      const dailyCount = countRes.rows[0].c;
      const warning = dailyCount >= MAX_ORDERS_PER_DAY ? `Note: ${dailyCount + 1} orders today, above the usual ${MAX_ORDERS_PER_DAY}/day for this BU.` : null;

      await client.query('BEGIN');
      const orderRef = await nextOrderRef();
      const orderRes = await client.query(
        `INSERT INTO horeca_orders (order_ref, salesman_user_id, salesperson_name, account_number, customer_name, site_use_id, location, po_number, warehouse, status)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,'pending') RETURNING id, order_ref, created_at`,
        [orderRef, req.user.uid, req.user.full_name || req.user.username, account_number, customer_name, site_use_id || null, location || null, po_number || null, warehouse || 'DCF']
      );
      const orderId = orderRes.rows[0].id;
      for (const l of lines) {
        await client.query(
          `INSERT INTO horeca_order_lines (order_id, item_code, description, qty, uom) VALUES ($1,$2,$3,$4,$5)`,
          [orderId, l.item_code, l.description || null, l.qty, String(l.uom).toUpperCase()]
        );
      }
      await client.query('COMMIT');
      res.json({ success: true, order_ref: orderRef, order_id: orderId, warning });
    } catch (e) {
      await client.query('ROLLBACK');
      res.status(500).json({ error: e.message });
    } finally {
      client.release();
    }
  });

  // Salesman: view own orders + status
  app.get('/api/horeca/orders/mine', requireAuth, requireHorecaRole('salesman'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT o.*, COALESCE(json_agg(json_build_object('item_code',l.item_code,'description',l.description,'qty',l.qty,'uom',l.uom)) FILTER (WHERE l.id IS NOT NULL), '[]') AS lines
         FROM horeca_orders o LEFT JOIN horeca_order_lines l ON l.order_id = o.id
         WHERE o.salesman_user_id = $1 GROUP BY o.id ORDER BY o.created_at DESC LIMIT 100`,
        [req.user.uid]
      );
      res.json({ orders: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Salesman: notifications
  app.get('/api/horeca/notifications', requireAuth, async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT n.*, o.order_ref FROM horeca_notifications n JOIN horeca_orders o ON o.id = n.order_id
         WHERE n.user_id = $1 ORDER BY n.created_at DESC LIMIT 50`,
        [req.user.uid]
      );
      res.json({ notifications: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  app.post('/api/horeca/notifications/:id/read', requireAuth, async function (req, res) {
    try {
      await pool.query(`UPDATE horeca_notifications SET read=true WHERE id=$1 AND user_id=$2`, [req.params.id, req.user.uid]);
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // CUSTOMER SERVICE: queue, approve, export Excel
  // ============================================================

  // Orders queue, grouped-friendly (returns flat list with salesperson_name; UI groups client-side)
  app.get('/api/horeca/orders', requireAuth, requireHorecaRole('cs', 'viewer', 'subadmin'), async function (req, res) {
    try {
      const status = req.query.status; // optional filter: pending / approved / sent
      let q = `SELECT o.*, COALESCE(json_agg(json_build_object('item_code',l.item_code,'description',l.description,'qty',l.qty,'uom',l.uom)) FILTER (WHERE l.id IS NOT NULL), '[]') AS lines
                FROM horeca_orders o LEFT JOIN horeca_order_lines l ON l.order_id = o.id`;
      const params = [];
      if (status) { params.push(status); q += ` WHERE o.status = $1`; }
      q += ` GROUP BY o.id ORDER BY o.salesperson_name, o.created_at DESC LIMIT 500`;
      const r = await pool.query(q, params);
      res.json({ orders: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Approve an order
  app.post('/api/horeca/orders/:id/approve', requireAuth, requireHorecaRole('cs', 'viewer', 'subadmin'), async function (req, res) {
    try {
      await pool.query(`UPDATE horeca_orders SET status='approved', actioned_at=NOW(), actioned_by=$1 WHERE id=$2`, [req.user.username, req.params.id]);
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Export Excel in the Doxtre/Oracle template format, then mark as sent + notify salesman
  app.get('/api/horeca/orders/:id/export-excel', requireAuth, requireHorecaRole('cs', 'viewer', 'subadmin'), async function (req, res) {
    try {
      const orderRes = await pool.query(`SELECT * FROM horeca_orders WHERE id=$1`, [req.params.id]);
      const order = orderRes.rows[0];
      if (!order) return res.status(404).json({ error: 'Order not found' });
      const linesRes = await pool.query(`SELECT * FROM horeca_order_lines WHERE order_id=$1 ORDER BY id`, [req.params.id]);

      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.columns = [
        { width: 20 }, // A Outer Barcode / Customer name labels
        { width: 22 }, // B EA Barcode / labels
        { width: 20 }, // C AKI Code / values
        { width: 34 }, // D Item Description
        { width: 14 }, // E SUB CATEGORY
        { width: 18 }, // F Price per pc/outer
        { width: 12 }, // G Order in PC
        { width: 10 }, // H FOC
        { width: 10 }, // I UOM
        { width: 10 }  // J TOTAL
      ];

      // Colors matched from the real template
      const FILL_ROW1   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0DFE6' } }; // light blue - title row
      const FILL_GENERAL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA1BFCD' } }; // medium blue - header block bg
      const FILL_LABEL   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4EA72E' } }; // green - key labels
      const FILL_YELLOW  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }; // yellow - W/H
      const FILL_ORANGE  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2AA84' } }; // orange - column headers
      const FILL_SUBCAT  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCFECF7' } }; // light blue - SUB CATEGORY
      const FILL_WHITE   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

      // Apply the general light-blue background across the whole header block first
      for (let r = 1; r <= 6; r++) {
        for (let c = 1; c <= 10; c++) {
          ws.getCell(r, c).fill = r <= 2 ? FILL_ROW1 : FILL_GENERAL;
          ws.getCell(r, c).font = { bold: true };
        }
      }

      ws.getCell('D1').value = 'ALPHAMED';
      ws.getCell('D2').value = 'HORECA ORDER FORM';
      ws.getCell('A3').value = 'Customer name :';
      ws.getCell('B3').value = 'Customer Nmae : '; ws.getCell('B3').fill = FILL_LABEL;
      ws.getCell('C3').value = order.customer_name;
      ws.getCell('F3').value = 'PO:';
      ws.getCell('G3').value = order.po_number || '';
      ws.getCell('A4').value = 'Customer Number  :';
      ws.getCell('B4').value = 'Cutomer code: '; ws.getCell('B4').fill = FILL_LABEL;
      const accountNumVal = order.account_number && !isNaN(order.account_number) ? Number(order.account_number) : order.account_number;
      ws.getCell('C4').value = accountNumVal; ws.getCell('C4').fill = FILL_WHITE; ws.getCell('C4').font = { bold: false };
      ws.getCell('F4').value = 'Sales person name : ' + (order.salesperson_name || '');
      ws.getCell('B5').value = 'W/H :'; ws.getCell('B5').fill = FILL_YELLOW;
      ws.getCell('C5').value = 'DCF';
      ws.getCell('B6').value = 'Location Site ID :'; ws.getCell('B6').fill = FILL_LABEL;
      const siteIdVal = order.site_use_id && !isNaN(order.site_use_id) ? Number(order.site_use_id) : (order.site_use_id || '');
      ws.getCell('C6').value = siteIdVal;
      ws.getCell('D6').value = 'Drp Down List';
      ws.getCell('F6').value = 'DATE: ' + new Date(order.created_at).toLocaleDateString('en-GB');

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

      linesRes.rows.forEach((l, idx) => {
        const r = headerRow + 1 + idx;
        const codeCell = ws.getCell(r, 3); codeCell.value = l.item_code; codeCell.font = { bold: false };       // C = AKI Code
        ws.getCell(r, 4).value = l.description || '';                                                            // D = Item Description
        const qtyCell = ws.getCell(r, 7); qtyCell.value = Number(l.qty); qtyCell.font = { bold: true }; qtyCell.alignment = { horizontal: 'center' }; // G = Order in PC
        ws.getCell(r, 9).value = l.uom;                                                                          // I = UOM
        const totalCell = ws.getCell(r, 10); totalCell.value = '-'; totalCell.font = { bold: true };             // J = TOTAL (Oracle-calculated, placeholder)
      });

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${order.order_ref}.xlsx"`);
      await wb.xlsx.write(res);

      const wasAlreadySent = order.status === 'sent';
      await pool.query(`UPDATE horeca_orders SET status='sent', actioned_at=NOW(), actioned_by=$1 WHERE id=$2`, [req.user.username, order.id]);
      if (!wasAlreadySent) {
        await pool.query(
          `INSERT INTO horeca_notifications (user_id, order_id, message) VALUES ($1,$2,$3)`,
          [order.salesman_user_id, order.id, `Your order ${order.order_ref} for ${order.customer_name} was sent to processing.`]
        );
      }
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  console.log('HoReCa module: routes mounted');

  // ============================================================
  // ADMIN: SALESMAN MANAGEMENT (create logins, assign to codes)
  // ============================================================

  // Distinct salesperson codes/names already present in the uploaded Customer Master
  // — used to populate a dropdown so admin picks the correct code instead of typing it.
  app.get('/api/horeca/salesperson-codes', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT DISTINCT salesperson_name FROM horeca_customer_sites
         WHERE salesperson_name IS NOT NULL AND salesperson_name != '' ORDER BY salesperson_name`
      );
      res.json({ codes: r.rows.map(row => ({ salesperson_name: row.salesperson_name })) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // List all salesman accounts, with count of customers/sites assigned to each code
  app.get('/api/horeca/salesmen', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT u.id, u.username, u.full_name, u.horeca_salesperson_name, u.active, u.last_login, u.created_at,
                (SELECT COUNT(*) FROM horeca_customer_sites s WHERE s.salesperson_name = u.horeca_salesperson_name) AS assigned_sites,
                (SELECT COUNT(*) FROM horeca_orders o WHERE o.salesman_user_id = u.id) AS total_orders
         FROM users u WHERE u.role = 'salesman' ORDER BY u.created_at DESC`
      );
      res.json({ salesmen: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Create a new salesman login
  app.post('/api/horeca/salesmen', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const { username, password, full_name, salesperson_name } = req.body;
      if (!username || !password || !full_name) return res.status(400).json({ error: 'Username, password and full name are required' });
      if (password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
      if (!salesperson_name) return res.status(400).json({ error: 'Please select the matching Salesperson Name from the Customer Master' });

      const existing = await pool.query('SELECT id FROM users WHERE username=$1', [username.toLowerCase().trim()]);
      if (existing.rows.length) return res.status(409).json({ error: 'Username already exists' });

      const hash = await bcrypt.hash(password, 10);
      const r = await pool.query(
        `INSERT INTO users (username, password_hash, full_name, role, horeca_salesperson_name, must_change_password, created_by)
         VALUES ($1,$2,$3,'salesman',$4,true,$5) RETURNING id, username, full_name, horeca_salesperson_name`,
        [username.toLowerCase().trim(), hash, full_name, salesperson_name, req.user.username]
      );
      await auditLog(req.user.uid, req.user.username, 'HORECA_SALESMAN_CREATE', `created ${username}`, '');
      res.json({ success: true, salesman: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Update a salesman (reassign code, activate/deactivate, rename)
  app.put('/api/horeca/salesmen/:id', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const { full_name, salesperson_name, active } = req.body;
      await pool.query(
        `UPDATE users SET
           full_name = COALESCE($1, full_name),
           horeca_salesperson_name = COALESCE($2, horeca_salesperson_name),
           active = COALESCE($3, active)
         WHERE id = $4 AND role = 'salesman'`,
        [full_name || null, salesperson_name || null, active, req.params.id]
      );
      await auditLog(req.user.uid, req.user.username, 'HORECA_SALESMAN_UPDATE', `updated user #${req.params.id}`, '');
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Reset a salesman's password
  app.post('/api/horeca/salesmen/:id/reset-password', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const { new_password } = req.body;
      if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
      const hash = await bcrypt.hash(new_password, 10);
      await pool.query(`UPDATE users SET password_hash=$1, must_change_password=true WHERE id=$2 AND role='salesman'`, [hash, req.params.id]);
      await auditLog(req.user.uid, req.user.username, 'HORECA_SALESMAN_RESET_PW', `reset password for user #${req.params.id}`, '');
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Delete a salesman account entirely
  app.delete('/api/horeca/salesmen/:id', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const check = await pool.query(`SELECT username FROM users WHERE id=$1 AND role='salesman'`, [req.params.id]);
      if (!check.rows.length) return res.status(404).json({ error: 'Salesman not found' });
      await pool.query(`DELETE FROM users WHERE id=$1 AND role='salesman'`, [req.params.id]);
      await auditLog(req.user.uid, req.user.username, 'HORECA_SALESMAN_DELETE', `deleted ${check.rows[0].username}`, '');
      res.json({ success: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
};
