// ============================================================
// VEHICLE MASTER MODULE — fleet registry used to fill in cost estimates
// when the dispatch file's own truck-type column is missing, by looking
// up the vehicle plate number instead.
// Mount with: require('./vehicle_master_module')(app, pool, requireAuth, requireRole, upload, auditLog, vehicleMasterMap);
// `vehicleMasterMap` is a plain object passed in by reference from server.js —
// this module populates it (never reassigns it) so server.js's synchronous
// dispatch-parsing code can read it without any DB calls mid-computation.
// ============================================================

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, vehicleMasterMap) {

  function normalizeVehicleNo(raw) {
    return String(raw || '').toUpperCase().replace(/\s+/g, '');
  }

  async function initVehicleMasterDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS vehicle_master (
        id SERIAL PRIMARY KEY,
        vehicle_no TEXT UNIQUE NOT NULL,
        vehicle_no_normalized TEXT UNIQUE NOT NULL,
        department TEXT,
        model TEXT,
        veh_type TEXT,
        vehicle_type_raw TEXT,
        weight TEXT,
        adhoc TEXT,
        partition_flag TEXT,
        active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_vehicle_master_norm ON vehicle_master(vehicle_no_normalized)`);
      console.log('Vehicle Master module: table ready');
      await refreshVehicleMasterMap();
    } catch (e) {
      console.error('Vehicle Master initDB error:', e.message);
    }
  }

  // Repopulate the shared in-memory map from the DB (mutates in place — never reassigns).
  async function refreshVehicleMasterMap() {
    try {
      const r = await pool.query(`SELECT vehicle_no_normalized, veh_type, vehicle_type_raw FROM vehicle_master WHERE active=true`);
      Object.keys(vehicleMasterMap).forEach(k => delete vehicleMasterMap[k]);
      r.rows.forEach(row => {
        vehicleMasterMap[row.vehicle_no_normalized] = { veh_type: row.veh_type, vehicle_type_raw: row.vehicle_type_raw };
      });
      console.log('Vehicle Master module: ' + r.rows.length + ' vehicles loaded into memory');
    } catch (e) {
      console.error('Vehicle Master map refresh error:', e.message);
    }
  }
  initVehicleMasterDB();

  // Upload the vehicle master Excel (add/update only, never delete — same pattern as HoReCa masters)
  app.post('/api/vehicle-master/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });

      let inserted = 0, updated = 0, skipped = 0;
      for (const row of rows) {
        const vehicleNo = String(row['Vehicle No.'] || row['Vehicle No'] || '').trim();
        if (!vehicleNo) { skipped++; continue; }
        const normalized = normalizeVehicleNo(vehicleNo);
        const department = row['Department'] || null;
        const model = row['Model'] || null;
        const vehType = row['Veh Type'] || null;
        const vehicleTypeRaw = row['Vehicle type'] || row['Vehicle Type'] || null;
        const weight = row['Weight'] || null;
        const adhoc = row['ADHOC'] || null;
        const partitionFlag = row['Partition'] || row['Partition '] || null;

        const existing = await pool.query(`SELECT id FROM vehicle_master WHERE vehicle_no_normalized=$1`, [normalized]);
        if (existing.rows.length) {
          await pool.query(
            `UPDATE vehicle_master SET vehicle_no=$1, department=$2, model=$3, veh_type=$4, vehicle_type_raw=$5, weight=$6, adhoc=$7, partition_flag=$8, updated_at=NOW()
             WHERE id=$9`,
            [vehicleNo, department, model, vehType, vehicleTypeRaw, weight, adhoc, partitionFlag, existing.rows[0].id]
          );
          updated++;
        } else {
          await pool.query(
            `INSERT INTO vehicle_master (vehicle_no, vehicle_no_normalized, department, model, veh_type, vehicle_type_raw, weight, adhoc, partition_flag)
             VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
            [vehicleNo, normalized, department, model, vehType, vehicleTypeRaw, weight, adhoc, partitionFlag]
          );
          inserted++;
        }
      }
      await refreshVehicleMasterMap();
      await auditLog(req.user.uid, req.user.username, 'VEHICLE_MASTER_UPLOAD', `inserted=${inserted} updated=${updated} skipped=${skipped}`, '');
      res.json({ success: true, inserted, updated, skipped, total_rows: rows.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Download what's currently stored, for verification
  app.get('/api/vehicle-master/export', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT vehicle_no AS "Vehicle No.", department AS "Department", model AS "Model", veh_type AS "Veh Type",
                vehicle_type_raw AS "Vehicle type", weight AS "Weight", adhoc AS "ADHOC", partition_flag AS "Partition"
         FROM vehicle_master WHERE active=true ORDER BY vehicle_no`
      );
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Vehicle Master');
      if (r.rows.length) {
        ws.columns = Object.keys(r.rows[0]).map(k => ({ header: k, key: k, width: 18 }));
        r.rows.forEach(row => ws.addRow(row));
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="Vehicle_Master_Current.xlsx"');
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      if (!res.headersSent) res.status(500).json({ error: e.message });
    }
  });

  console.log('Vehicle Master module: routes mounted');
};
