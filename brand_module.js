// ============================================================
// BRAND PRESENTATIONS MODULE — plugs into existing AZHAR-AI server.js
// Self-contained: own tables, shares the existing pg pool + auth.
// Mount with: require('./brand_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);
//
// Access model:
//   - superadmin / subadmin: see and manage every brand, no grant needed.
//   - anyone else: needs a row in brand_user_access for that brand.
//       role='manager' -> can edit sections, upload Excel, upload documents
//       role='viewer'  -> read-only
//
// File storage note: this codebase never persists raw uploaded files to
// disk (Render's filesystem is ephemeral, and every other module only
// keeps the *extracted* data from an Excel buffer). Brand documents
// (catalogs, price lists, videos) are different — people need to actually
// download the original file later, not just its extracted numbers. With
// no S3/object storage configured anywhere in this app yet, this module
// stores document bytes directly in Postgres (BYTEA) so it works today
// with zero new infrastructure. Fine for catalogs/PDFs; if brand videos
// start running into the tens of MB regularly, that's the point to add
// real object storage (S3, Cloudinary, a Render persistent disk) instead —
// flagging that now rather than silently building around it.
// ============================================================

const XLSX = require('xlsx');

module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, bcrypt) {

  // Section types an Excel upload is allowed to overwrite.
  // Everything else (swot_radar, initiative_grid, cover, document_list) is
  // manual-only and an Excel refresh must never touch it.
  const EXCEL_FILLABLE_TYPES = ['kpi_hero', 'sku_rings', 'stock_table', 'ranked_table'];

  // ── Init tables ──
  async function initBrandsDB() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS brand_template_types (
        key                 TEXT PRIMARY KEY,
        display_name        TEXT NOT NULL,
        description         TEXT,
        is_excel_fillable   BOOLEAN NOT NULL DEFAULT false,
        data_schema_hint    JSONB,
        created_at          TIMESTAMPTZ DEFAULT NOW()
      )`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brands (
        id              SERIAL PRIMARY KEY,
        name            TEXT NOT NULL,
        slug            TEXT UNIQUE NOT NULL,
        accent_color    TEXT NOT NULL DEFAULT '#c6a15b',
        logo_url        TEXT,
        tagline         TEXT,
        status          TEXT NOT NULL DEFAULT 'draft',
        created_by      INT REFERENCES users(id),
        created_at      TIMESTAMPTZ DEFAULT NOW(),
        updated_at      TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brands_status ON brands(status)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brand_sections (
        id                 SERIAL PRIMARY KEY,
        brand_id           INT REFERENCES brands(id) ON DELETE CASCADE,
        template_type      TEXT NOT NULL,
        section_order      INT NOT NULL DEFAULT 0,
        chapter_tag        TEXT,
        title              TEXT,
        data               JSONB NOT NULL DEFAULT '{}'::jsonb,
        data_source        TEXT NOT NULL DEFAULT 'manual',
        last_refreshed_at  TIMESTAMPTZ,
        updated_by         INT REFERENCES users(id),
        created_at         TIMESTAMPTZ DEFAULT NOW(),
        updated_at         TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(brand_id, section_order)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_sections_lookup ON brand_sections(brand_id, template_type)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brand_section_history (
        id            SERIAL PRIMARY KEY,
        section_id    INT REFERENCES brand_sections(id) ON DELETE CASCADE,
        data          JSONB NOT NULL,
        data_source   TEXT NOT NULL,
        replaced_by   INT REFERENCES users(id),
        replaced_at   TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_section_history_section ON brand_section_history(section_id, replaced_at DESC)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brand_documents (
        id              SERIAL PRIMARY KEY,
        brand_id        INT REFERENCES brands(id) ON DELETE CASCADE,
        title           TEXT NOT NULL,
        file_data       BYTEA NOT NULL,
        mime_type       TEXT,
        file_size_bytes BIGINT,
        note            TEXT,
        uploaded_by     INT REFERENCES users(id),
        uploaded_at     TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_documents_brand ON brand_documents(brand_id)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brand_user_access (
        id           SERIAL PRIMARY KEY,
        brand_id     INT REFERENCES brands(id) ON DELETE CASCADE,
        user_id      INT REFERENCES users(id) ON DELETE CASCADE,
        role         TEXT NOT NULL DEFAULT 'viewer',
        granted_by   INT REFERENCES users(id),
        granted_at   TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(brand_id, user_id)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_access_user ON brand_user_access(user_id)`);

      await pool.query(`CREATE TABLE IF NOT EXISTS brand_excel_uploads (
        id                SERIAL PRIMARY KEY,
        brand_id          INT REFERENCES brands(id) ON DELETE CASCADE,
        file_name         TEXT NOT NULL,
        uploaded_by       INT REFERENCES users(id),
        status            TEXT NOT NULL DEFAULT 'success',
        sections_updated  TEXT[],
        error_message     TEXT,
        uploaded_at       TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_excel_uploads_brand ON brand_excel_uploads(brand_id, uploaded_at DESC)`);

      // Dated, browsable archive: a full point-in-time copy of every section's data.
      // One row per Excel refresh (auto) or manual "Save version" (brand manager).
      // This is what lets someone open "Little Clouds — July 2026" later, distinct
      // from brand_section_history which is per-section audit trail, not a
      // browsable whole-deck snapshot.
      await pool.query(`CREATE TABLE IF NOT EXISTS brand_snapshots (
        id              SERIAL PRIMARY KEY,
        brand_id        INT REFERENCES brands(id) ON DELETE CASCADE,
        label           TEXT,
        sections_data   JSONB NOT NULL,
        source          TEXT NOT NULL DEFAULT 'excel_upload',
        created_by      INT REFERENCES users(id),
        created_at      TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_snapshots_brand ON brand_snapshots(brand_id, created_at DESC)`);

      // Seed template type reference rows (idempotent — ON CONFLICT DO NOTHING)
      const templateTypes = [
        ['cover', 'Cover', 'Brand mark, tagline, live/sample status badge.', false],
        ['kpi_hero', 'KPI Hero Numbers', 'Large animated count-up KPIs.', true],
        ['sku_rings', 'SKU Performance Rings', 'Radial progress rings — achieved vs target per SKU.', true],
        ['stock_table', 'Stock Status Table', 'SKU-wise value, volume, ageing, shelf-life.', true],
        ['ranked_table', 'Ranked Performance Table', 'Generic ranked table with growth% and contribution%.', true],
        ['swot_radar', 'SWOT Radar', 'Interactive radar chart with detail panels.', false],
        ['initiative_grid', 'Initiative Cards', 'Grid of planned/ongoing activities with status.', false],
        ['document_list', 'Documents', 'Downloadable files — backed by brand_documents.', false],
      ];
      for (const [key, name, desc, fillable] of templateTypes) {
        await pool.query(
          `INSERT INTO brand_template_types (key, display_name, description, is_excel_fillable)
           VALUES ($1,$2,$3,$4) ON CONFLICT (key) DO NOTHING`,
          [key, name, desc, fillable]
        );
      }

      console.log('Brand module: tables ready');
    } catch (e) {
      console.error('Brand module initDB error:', e.message);
    }
  }
  initBrandsDB();

  // ── Helpers ──
  const ADMIN_ROLES = ['superadmin', 'subadmin'];

  function isAdmin(user) { return ADMIN_ROLES.includes(user.role); }

  // Loads the caller's access level for one brand: 'admin' | 'manager' | 'viewer' | null
  async function getAccessLevel(user, brandId) {
    if (isAdmin(user)) return 'admin';
    const r = await pool.query(
      `SELECT role FROM brand_user_access WHERE brand_id=$1 AND user_id=$2`,
      [brandId, user.uid]
    );
    return r.rows[0] ? r.rows[0].role : null;
  }

  // Middleware: caller must at least be able to view this brand (:id or :brandId param)
  function requireBrandView() {
    return async function (req, res, next) {
      const brandId = req.params.id || req.params.brandId;
      const level = await getAccessLevel(req.user, brandId);
      if (!level) return res.status(403).json({ error: 'No access to this brand' });
      req.brandAccessLevel = level;
      next();
    };
  }

  // Middleware: caller must be able to edit this brand (admin or manager)
  function requireBrandManage() {
    return async function (req, res, next) {
      const brandId = req.params.id || req.params.brandId;
      const level = await getAccessLevel(req.user, brandId);
      if (level !== 'admin' && level !== 'manager') {
        return res.status(403).json({ error: 'Manager access required for this brand' });
      }
      req.brandAccessLevel = level;
      next();
    };
  }

  function slugify(name) {
    return String(name).toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }

  // ============================================================
  // BRANDS — list / create / read / update
  // ============================================================

  // List brands visible to the caller (all for admin, granted-only otherwise)
  app.get('/api/brands', requireAuth, async function (req, res) {
    try {
      let rows;
      if (isAdmin(req.user)) {
        const r = await pool.query(`
          SELECT b.*,
            (SELECT string_agg(DISTINCT COALESCE(u.full_name, u.username), ', ')
             FROM brand_user_access a JOIN users u ON u.id = a.user_id
             WHERE a.brand_id = b.id AND a.role = 'manager') AS managers,
            (SELECT MAX(s.updated_at) FROM brand_sections s WHERE s.brand_id = b.id) AS content_updated_at
          FROM brands b ORDER BY b.name`);
        rows = r.rows;
      } else {
        const r = await pool.query(
          `SELECT b.*, a.role AS my_role FROM brands b
           JOIN brand_user_access a ON a.brand_id = b.id
           WHERE a.user_id = $1 ORDER BY b.name`,
          [req.user.uid]
        );
        rows = r.rows;
      }
      res.json({ brands: rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Create a brand — self-serve. Anyone with a login except 'salesman' (a fixed,
  // separate order-taking flow) can create their own brand and becomes its
  // manager automatically. Admins never need to manually assign this — they see
  // every brand regardless (isAdmin() bypass), purely for oversight.
  app.post('/api/brands', requireAuth, async function (req, res) {
    const client = await pool.connect();
    try {
      if (req.user.role === 'salesman') {
        return res.status(403).json({ error: 'Salesman logins are for HoReCa order-taking only' });
      }
      const { name, accent_color, logo_url, tagline } = req.body;
      if (!name) return res.status(400).json({ error: 'Brand name is required' });
      const slug = slugify(name);

      await client.query('BEGIN');
      const r = await client.query(
        `INSERT INTO brands (name, slug, accent_color, logo_url, tagline, status, created_by)
         VALUES ($1,$2,$3,$4,$5,'draft',$6) RETURNING *`,
        [name, slug, accent_color || '#c6a15b', logo_url || null, tagline || null, req.user.uid]
      );
      const brand = r.rows[0];

      // Default cover section — without this the deck opens to a blank screen
      // until the first Excel upload or manual edit.
      await client.query(
        `INSERT INTO brand_sections (brand_id, template_type, section_order, chapter_tag, title, data, data_source, updated_by)
         VALUES ($1,'cover',0,NULL,NULL,$2,'manual',$3)`,
        [brand.id, JSON.stringify({ tagline: tagline || '' }), req.user.uid]
      );

      // Creator becomes manager automatically — this IS the assignment, no
      // separate admin action needed. Admins don't need a row here (bypass),
      // but non-admin creators do, or they'd have no access to what they just made.
      if (!isAdmin(req.user)) {
        await client.query(
          `INSERT INTO brand_user_access (brand_id, user_id, role, granted_by) VALUES ($1,$2,'manager',$2)`,
          [brand.id, req.user.uid]
        );
      }

      await client.query('COMMIT');
      await auditLog(req.user.uid, req.user.username, 'brand_created', name, req.ip);
      res.json({ brand: brand });
    } catch (e) {
      await client.query('ROLLBACK');
      if (e.code === '23505') return res.status(409).json({ error: 'A brand with that name already exists' });
      res.status(500).json({ error: e.message });
    } finally {
      client.release();
    }
  });

  // Full brand detail — meta + sections + document list (view access required)
  app.get('/api/brands/:id', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const brandR = await pool.query(`SELECT * FROM brands WHERE id=$1`, [req.params.id]);
      if (!brandR.rows[0]) return res.status(404).json({ error: 'Brand not found' });

      const sectionsR = await pool.query(
        `SELECT id, template_type, section_order, chapter_tag, title, data, data_source, last_refreshed_at
         FROM brand_sections WHERE brand_id=$1 ORDER BY section_order`,
        [req.params.id]
      );
      const docsR = await pool.query(
        `SELECT id, title, mime_type, file_size_bytes, note, uploaded_at FROM brand_documents
         WHERE brand_id=$1 ORDER BY uploaded_at DESC`,
        [req.params.id]
      );

      res.json({
        brand: brandR.rows[0],
        sections: sectionsR.rows,
        documents: docsR.rows,
        access_level: req.brandAccessLevel
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Update brand meta (admin only — changing accent color/logo/status is brand-level, not section-level)
  app.put('/api/brands/:id', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const { name, accent_color, logo_url, tagline, status } = req.body;
      const r = await pool.query(
        `UPDATE brands SET
           name = COALESCE($1, name),
           accent_color = COALESCE($2, accent_color),
           logo_url = COALESCE($3, logo_url),
           tagline = COALESCE($4, tagline),
           status = COALESCE($5, status),
           updated_at = NOW()
         WHERE id=$6 RETURNING *`,
        [name || null, accent_color || null, logo_url || null, tagline || null, status || null, req.params.id]
      );
      if (!r.rows[0]) return res.status(404).json({ error: 'Brand not found' });
      res.json({ brand: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // SECTIONS — manual create/update (brand manager or admin)
  // ============================================================

  // Upsert one section. If section_order already exists for this brand, updates it
  // (and snapshots the old data into brand_section_history first).
  app.post('/api/brands/:id/sections', requireAuth, requireBrandManage(), async function (req, res) {
    const client = await pool.connect();
    try {
      const brandId = req.params.id;
      const { template_type, section_order, chapter_tag, title, data } = req.body;
      if (!template_type || section_order == null) {
        return res.status(400).json({ error: 'template_type and section_order are required' });
      }

      await client.query('BEGIN');

      const existing = await client.query(
        `SELECT id, data, data_source FROM brand_sections WHERE brand_id=$1 AND section_order=$2`,
        [brandId, section_order]
      );

      let section;
      if (existing.rows[0]) {
        await client.query(
          `INSERT INTO brand_section_history (section_id, data, data_source, replaced_by)
           VALUES ($1,$2,$3,$4)`,
          [existing.rows[0].id, existing.rows[0].data, existing.rows[0].data_source, req.user.uid]
        );
        const r = await client.query(
          `UPDATE brand_sections SET
             template_type=$1, chapter_tag=$2, title=$3, data=$4,
             data_source='manual', updated_by=$5, updated_at=NOW()
           WHERE id=$6 RETURNING *`,
          [template_type, chapter_tag || null, title || null, JSON.stringify(data || {}), req.user.uid, existing.rows[0].id]
        );
        section = r.rows[0];
      } else {
        const r = await client.query(
          `INSERT INTO brand_sections (brand_id, template_type, section_order, chapter_tag, title, data, data_source, updated_by)
           VALUES ($1,$2,$3,$4,$5,$6,'manual',$7) RETURNING *`,
          [brandId, template_type, section_order, chapter_tag || null, title || null, JSON.stringify(data || {}), req.user.uid]
        );
        section = r.rows[0];
      }

      await client.query('COMMIT');
      await auditLog(req.user.uid, req.user.username, 'brand_section_updated', `brand ${brandId} section ${section_order} (${template_type})`, req.ip);
      res.json({ section });
    } catch (e) {
      await client.query('ROLLBACK');
      res.status(500).json({ error: e.message });
    } finally {
      client.release();
    }
  });

  // Delete a section
  app.delete('/api/brands/:id/sections/:sectionId', requireAuth, requireBrandManage(), async function (req, res) {
    try {
      await pool.query(`DELETE FROM brand_sections WHERE id=$1 AND brand_id=$2`, [req.params.sectionId, req.params.id]);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // EXCEL UPLOAD — refreshes only excel-fillable sections
  // ============================================================
  // Expected tabs (matches Little_Clouds_Data_Template.xlsx from the prototype):
  //   "Target vs Achievement"   -> kpi_hero + sku_rings
  //   "Distribution Plan"       -> kpi_hero (store coverage)
  //   "Stock Status"            -> stock_table
  //   "Sellout Noon" / any ranked metric sheet -> ranked_table
  // Column headers are read starting at row 4 (range:3), matching the template's
  // title/legend rows above the header row.

  app.post('/api/brands/:id/upload-excel', requireAuth, requireBrandManage(), upload.single('file'), async function (req, res) {
    const client = await pool.connect();
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const brandId = req.params.id;
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });

      const sheetJson = (name) => {
        if (!wb.Sheets[name]) return [];
        return XLSX.utils.sheet_to_json(wb.Sheets[name], { range: 3, defval: '' });
      };

      const sectionsUpdated = [];
      await client.query('BEGIN');

      async function upsertExcelSection(templateType, sectionOrder, chapterTag, title, data) {
        const existing = await client.query(
          `SELECT id, data, data_source FROM brand_sections WHERE brand_id=$1 AND template_type=$2 AND section_order=$3`,
          [brandId, templateType, sectionOrder]
        );
        if (existing.rows[0]) {
          await client.query(
            `INSERT INTO brand_section_history (section_id, data, data_source, replaced_by) VALUES ($1,$2,$3,$4)`,
            [existing.rows[0].id, existing.rows[0].data, existing.rows[0].data_source, req.user.uid]
          );
          await client.query(
            `UPDATE brand_sections SET data=$1, data_source='excel_upload', last_refreshed_at=NOW(), updated_by=$2, updated_at=NOW()
             WHERE id=$3`,
            [JSON.stringify(data), req.user.uid, existing.rows[0].id]
          );
        } else {
          await client.query(
            `INSERT INTO brand_sections (brand_id, template_type, section_order, chapter_tag, title, data, data_source, last_refreshed_at, updated_by)
             VALUES ($1,$2,$3,$4,$5,$6,'excel_upload',NOW(),$7)`,
            [brandId, templateType, sectionOrder, chapterTag, title, JSON.stringify(data), req.user.uid]
          );
        }
        sectionsUpdated.push(templateType);
      }

      // --- Target vs Achievement -> sku_rings + kpi_hero ---
      const targetRows = sheetJson('Target vs Achievement').filter(r => r['SKU Name']);
      if (targetRows.length) {
        const items = targetRows.map(r => ({
          sku: r['SKU Name'],
          target: Number(r['Target (Units)']) || 0,
          achieved: Number(r['Achieved (Units)']) || 0
        }));
        await upsertExcelSection('sku_rings', 2, '01 — Sales Performance · IMS', 'Top SKU performance', { items });

        const totalTarget = items.reduce((a, b) => a + b.target, 0);
        const totalAchieved = items.reduce((a, b) => a + b.achieved, 0);
        await upsertExcelSection('kpi_hero', 1, '01 — Sales Performance · IMS', 'Target vs. Achievement', {
          kpis: [{ label: 'Achieved / Target', value: totalAchieved, compare_value: totalTarget, unit: 'units' }]
        });
      }

      // --- Distribution Plan -> folds into kpi_hero ---
      const distRows = sheetJson('Distribution Plan').filter(r => r['Region / Channel']);
      if (distRows.length) {
        const planned = distRows.reduce((a, r) => a + (Number(r['Planned Stores']) || 0), 0);
        const actual = distRows.reduce((a, r) => a + (Number(r['Actual Stores Reached']) || 0), 0);
        const existingKpi = await client.query(
          `SELECT data FROM brand_sections WHERE brand_id=$1 AND template_type='kpi_hero' AND section_order=1`,
          [brandId]
        );
        const kpis = existingKpi.rows[0] ? existingKpi.rows[0].data.kpis || [] : [];
        kpis.push({ label: 'Store Coverage', value: actual, compare_value: planned, unit: 'stores' });
        await upsertExcelSection('kpi_hero', 1, '01 — Sales Performance · IMS', 'Target vs. Achievement', { kpis });
      }

      // --- Stock Status -> stock_table ---
      const stockRows = sheetJson('Stock Status').filter(r => r['SKU Name']);
      if (stockRows.length) {
        const rows = stockRows.map(r => ({
          sku: r['SKU Name'],
          value: Number(r['Stock Value (AED)']) || 0,
          volume: Number(r['Volume (Units)']) || 0,
          ageing: r['Ageing Bucket (0-30/31-60/61+)'] || '0-30',
          shelf_status: r['Shelf Life Status (Healthy/Monitor/Near Expiry)'] || 'Healthy'
        }));
        await upsertExcelSection('stock_table', 3, '03 — SKU-wise Stock Status', 'Value, ageing & shelf life', { rows });
      }

      // --- Sellout Noon -> kpi_hero (its own section slot) ---
      const selloutRows = sheetJson('Sellout Noon').filter(r => r['Metric']);
      if (selloutRows.length) {
        const find = (m) => {
          const row = selloutRows.find(r => String(r['Metric']).toLowerCase().includes(m));
          return row ? Number(row['This Month']) || 0 : null;
        };
        const kpis = [];
        const u = find('sellout units'); if (u !== null) kpis.push({ label: 'Sellout Units MTD', value: u, unit: 'units' });
        const rv = find('sellout revenue'); if (rv !== null) kpis.push({ label: 'Sellout Revenue', value: rv, unit: 'AED' });
        const cv = find('conversion'); if (cv !== null) kpis.push({ label: 'Conversion Rate', value: cv, unit: '%' });
        const sc = find('stock cover'); if (sc !== null) kpis.push({ label: 'Stock Cover', value: sc, unit: 'days' });
        if (kpis.length) await upsertExcelSection('kpi_hero', 5, '02 — Sales Performance · Sellout (Noon)', 'Noon.com channel', { kpis });
      }

      if (sectionsUpdated.length === 0) {
        await client.query('ROLLBACK');
        await pool.query(
          `INSERT INTO brand_excel_uploads (brand_id, file_name, uploaded_by, status, error_message)
           VALUES ($1,$2,$3,'failed',$4)`,
          [brandId, req.file.originalname, req.user.uid, 'No recognized tabs found — check tab names match the template exactly']
        );
        return res.status(400).json({ error: "Couldn't find any recognized tabs. Check the file matches the template (tab names must match exactly)." });
      }

      await client.query('COMMIT');
      await pool.query(
        `INSERT INTO brand_excel_uploads (brand_id, file_name, uploaded_by, status, sections_updated)
         VALUES ($1,$2,$3,'success',$4)`,
        [brandId, req.file.originalname, req.user.uid, sectionsUpdated]
      );

      // Dated archive entry — a full copy of every section as it stands right
      // now, so this exact moment can be reopened later from the history browser.
      const allSections = await pool.query(
        `SELECT template_type, section_order, chapter_tag, title, data FROM brand_sections WHERE brand_id=$1 ORDER BY section_order`,
        [brandId]
      );
      const label = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      await pool.query(
        `INSERT INTO brand_snapshots (brand_id, label, sections_data, source, created_by) VALUES ($1,$2,$3,'excel_upload',$4)`,
        [brandId, label, JSON.stringify(allSections.rows), req.user.uid]
      );

      await auditLog(req.user.uid, req.user.username, 'brand_excel_refresh', `brand ${brandId}: ${sectionsUpdated.join(', ')}`, req.ip);
      res.json({ ok: true, sections_updated: sectionsUpdated });
    } catch (e) {
      await client.query('ROLLBACK');
      res.status(500).json({ error: e.message });
    } finally {
      client.release();
    }
  });

  // ============================================================
  // SNAPSHOTS — dated, browsable archive (brand -> year -> date)
  // ============================================================

  app.get('/api/brands/:id/snapshots', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT s.id, s.label, s.source, s.created_at, u.full_name AS created_by_name
         FROM brand_snapshots s LEFT JOIN users u ON u.id = s.created_by
         WHERE s.brand_id = $1 ORDER BY s.created_at DESC`,
        [req.params.id]
      );
      res.json({ snapshots: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.get('/api/brands/:id/snapshots/:snapshotId', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT id, label, sections_data, created_at FROM brand_snapshots WHERE id=$1 AND brand_id=$2`,
        [req.params.snapshotId, req.params.id]
      );
      if (!r.rows[0]) return res.status(404).json({ error: 'Snapshot not found' });
      res.json({ snapshot: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Manual "save version" — for brand managers editing by hand (not Excel),
  // so their dated archive isn't only ever populated by uploads.
  app.post('/api/brands/:id/snapshots', requireAuth, requireBrandManage(), async function (req, res) {
    try {
      const allSections = await pool.query(
        `SELECT template_type, section_order, chapter_tag, title, data FROM brand_sections WHERE brand_id=$1 ORDER BY section_order`,
        [req.params.id]
      );
      const label = req.body.label || new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      const r = await pool.query(
        `INSERT INTO brand_snapshots (brand_id, label, sections_data, source, created_by) VALUES ($1,$2,$3,'manual',$4) RETURNING id, label, created_at`,
        [req.params.id, label, JSON.stringify(allSections.rows), req.user.uid]
      );
      res.json({ snapshot: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // DOCUMENTS — catalogs, price lists, videos (stored as BYTEA — see note at top)
  // ============================================================

  app.post('/api/brands/:id/documents', requireAuth, requireBrandManage(), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const { title, note } = req.body;
      const r = await pool.query(
        `INSERT INTO brand_documents (brand_id, title, file_data, mime_type, file_size_bytes, note, uploaded_by)
         VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id, title, mime_type, file_size_bytes, note, uploaded_at`,
        [req.params.id, title || req.file.originalname, req.file.buffer, req.file.mimetype, req.file.size, note || null, req.user.uid]
      );
      await auditLog(req.user.uid, req.user.username, 'brand_document_uploaded', `brand ${req.params.id}: ${title || req.file.originalname}`, req.ip);
      res.json({ document: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.get('/api/brands/:id/documents/:docId/download', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT title, file_data, mime_type FROM brand_documents WHERE id=$1 AND brand_id=$2`,
        [req.params.docId, req.params.id]
      );
      if (!r.rows[0]) return res.status(404).json({ error: 'Document not found' });
      const doc = r.rows[0];
      res.setHeader('Content-Type', doc.mime_type || 'application/octet-stream');
      res.setHeader('Content-Disposition', `attachment; filename="${doc.title}"`);
      res.send(doc.file_data);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.delete('/api/brands/:id/documents/:docId', requireAuth, requireBrandManage(), async function (req, res) {
    try {
      await pool.query(`DELETE FROM brand_documents WHERE id=$1 AND brand_id=$2`, [req.params.docId, req.params.id]);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // ACCESS CONTROL — who can see/edit which brand (admin only)
  // ============================================================

  app.get('/api/brands/:id/access', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT a.id, a.role, a.granted_at, u.id AS user_id, u.username, u.full_name
         FROM brand_user_access a JOIN users u ON u.id = a.user_id
         WHERE a.brand_id = $1 ORDER BY u.full_name`,
        [req.params.id]
      );
      res.json({ access: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.post('/api/brands/:id/access', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      const { user_id, role } = req.body;
      if (!user_id || !['manager', 'viewer'].includes(role)) {
        return res.status(400).json({ error: 'user_id and role (manager|viewer) are required' });
      }
      const r = await pool.query(
        `INSERT INTO brand_user_access (brand_id, user_id, role, granted_by)
         VALUES ($1,$2,$3,$4)
         ON CONFLICT (brand_id, user_id) DO UPDATE SET role=$3, granted_by=$4, granted_at=NOW()
         RETURNING *`,
        [req.params.id, user_id, role, req.user.uid]
      );
      await auditLog(req.user.uid, req.user.username, 'brand_access_granted', `brand ${req.params.id} -> user ${user_id} (${role})`, req.ip);
      res.json({ access: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.delete('/api/brands/:id/access/:userId', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      await pool.query(`DELETE FROM brand_user_access WHERE brand_id=$1 AND user_id=$2`, [req.params.id, req.params.userId]);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  console.log('Brand module: routes mounted');
};
