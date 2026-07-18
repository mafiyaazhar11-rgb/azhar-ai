// ============================================================
// BRAND PRESENTATIONS MODULE — plugs into existing AZHAR-AI server.js
// Self-contained: own tables, shares the existing pg pool + auth.
// Mount with: require('./brand_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);
//
// Access model:
//   - superadmin ONLY: sees and manages every brand, no grant needed. Deliberately
//     narrower than the rest of the app (which often treats subadmin as a near-
//     admin uploader role) — brand oversight is superadmin-specific, per instruction.
//   - anyone else (including subadmin): needs a row in brand_user_access for that brand.
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
        logo_data       BYTEA,
        logo_mime       TEXT,
        tagline         TEXT,
        status          TEXT NOT NULL DEFAULT 'draft',
        created_by      INT REFERENCES users(id),
        created_at      TIMESTAMPTZ DEFAULT NOW(),
        updated_at      TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`ALTER TABLE brands ADD COLUMN IF NOT EXISTS logo_data BYTEA`);
      await pool.query(`ALTER TABLE brands ADD COLUMN IF NOT EXISTS logo_mime TEXT`);
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

      // ── Phase 1: proper draft -> validate -> preview -> publish versioning ──
      // Supersedes brand_snapshots for anything Excel-driven. brand_snapshots stays
      // in the schema untouched (no data loss, nothing that reads it breaks) but new
      // uploads flow through here instead. A version is a full point-in-time copy of
      // every section, same as a snapshot, plus the metadata a real leadership
      // presentation needs: a name, a reporting period, a status, and what validation
      // found. Publishing one supersedes the previous published version — it is never
      // deleted, only marked superseded, so "previous version available" always holds.
      await pool.query(`CREATE TABLE IF NOT EXISTS brand_versions (
        id                 SERIAL PRIMARY KEY,
        brand_id           INT REFERENCES brands(id) ON DELETE CASCADE,
        version_number     INT NOT NULL,
        presentation_name  TEXT NOT NULL,
        reporting_month    INT,
        reporting_year     INT,
        status             TEXT NOT NULL DEFAULT 'draft',
        sections_data      JSONB NOT NULL,
        validation         JSONB NOT NULL DEFAULT '{"errors":[],"warnings":[]}'::jsonb,
        source_file_name   TEXT,
        notes              TEXT,
        created_by         INT REFERENCES users(id),
        created_at         TIMESTAMPTZ DEFAULT NOW(),
        published_by       INT REFERENCES users(id),
        published_at       TIMESTAMPTZ,
        UNIQUE(brand_id, version_number)
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_versions_brand ON brand_versions(brand_id, version_number DESC)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_brand_versions_status ON brand_versions(brand_id, status)`);

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
  const ADMIN_ROLES = ['superadmin']; // deliberately excludes subadmin — see note above

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
  // VALIDATION ENGINE — row-level checks on the parsed Excel data.
  // Returns { errors: [...], warnings: [...] }. Each item:
  //   { sheet, row, column, message }
  // `row` is the actual spreadsheet row number (matches what the manager
  // sees in Excel — row 5 is the first data row, since row 4 is the header).
  // Errors BLOCK publish. Warnings don't, but are shown in the preview.
  // ============================================================
  function isBlank(v) { return v === null || v === undefined || String(v).trim() === ''; }
  function isNumeric(v) { return v !== '' && v !== null && v !== undefined && !isNaN(Number(v)); }

  function validateTargetVsAchievement(rows, errors, warnings) {
    const seen = new Set();
    rows.forEach((r, i) => {
      const excelRow = i + 5;
      const sku = r['SKU Name'];
      if (isBlank(sku)) { errors.push({ sheet: 'Target vs Achievement', row: excelRow, column: 'SKU Name', message: 'SKU Name is missing' }); return; }
      const key = String(sku).trim().toLowerCase();
      if (seen.has(key)) errors.push({ sheet: 'Target vs Achievement', row: excelRow, column: 'SKU Name', message: `Duplicate SKU "${sku}" — already appears earlier in this sheet` });
      seen.add(key);
      if (!isNumeric(r['Target (Units)'])) errors.push({ sheet: 'Target vs Achievement', row: excelRow, column: 'Target (Units)', message: `"${r['Target (Units)']}" is not a valid number` });
      if (!isNumeric(r['Achieved (Units)'])) errors.push({ sheet: 'Target vs Achievement', row: excelRow, column: 'Achieved (Units)', message: `"${r['Achieved (Units)']}" is not a valid number` });
      if (isNumeric(r['Target (Units)']) && Number(r['Target (Units)']) < 0) warnings.push({ sheet: 'Target vs Achievement', row: excelRow, column: 'Target (Units)', message: 'Negative target — please confirm this is intentional' });
    });
  }
  function validateDistributionPlan(rows, errors, warnings) {
    const seen = new Set();
    rows.forEach((r, i) => {
      const excelRow = i + 5;
      const region = r['Region / Channel'];
      if (isBlank(region)) { warnings.push({ sheet: 'Distribution Plan', row: excelRow, column: 'Region / Channel', message: 'Row skipped — no region/channel given' }); return; }
      const key = String(region).trim().toLowerCase();
      if (seen.has(key)) warnings.push({ sheet: 'Distribution Plan', row: excelRow, column: 'Region / Channel', message: `Duplicate region "${region}" — values will be summed together` });
      seen.add(key);
      if (!isNumeric(r['Planned Stores'])) errors.push({ sheet: 'Distribution Plan', row: excelRow, column: 'Planned Stores', message: `"${r['Planned Stores']}" is not a valid number` });
      if (!isNumeric(r['Actual Stores Reached'])) errors.push({ sheet: 'Distribution Plan', row: excelRow, column: 'Actual Stores Reached', message: `"${r['Actual Stores Reached']}" is not a valid number` });
    });
  }
  function validateStockStatus(rows, errors, warnings) {
    const AGEING_VALUES = ['0-30', '31-60', '61+'];
    const SHELF_VALUES = ['Healthy', 'Monitor', 'Near Expiry'];
    const seen = new Set();
    rows.forEach((r, i) => {
      const excelRow = i + 5;
      const sku = r['SKU Name'];
      if (isBlank(sku)) { errors.push({ sheet: 'Stock Status', row: excelRow, column: 'SKU Name', message: 'SKU Name is missing' }); return; }
      const key = String(sku).trim().toLowerCase();
      if (seen.has(key)) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'SKU Name', message: `Duplicate SKU "${sku}" — already appears earlier in this sheet` });
      seen.add(key);
      if (!isNumeric(r['Stock Value (AED)'])) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Stock Value (AED)', message: `"${r['Stock Value (AED)']}" is not a valid number` });
      if (!isNumeric(r['Volume (Units)'])) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Volume (Units)', message: `"${r['Volume (Units)']}" is not a valid number` });
      const ageing = r['Ageing Bucket (0-30/31-60/61+)'];
      if (isBlank(ageing)) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Ageing Bucket', message: 'Ageing bucket is missing' });
      else if (!AGEING_VALUES.includes(String(ageing).trim())) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Ageing Bucket', message: `"${ageing}" is not one of: ${AGEING_VALUES.join(', ')}` });
      const shelf = r['Shelf Life Status (Healthy/Monitor/Near Expiry)'];
      if (isBlank(shelf)) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Shelf Life Status', message: 'Shelf life status is missing' });
      else if (!SHELF_VALUES.includes(String(shelf).trim())) errors.push({ sheet: 'Stock Status', row: excelRow, column: 'Shelf Life Status', message: `"${shelf}" is not one of: ${SHELF_VALUES.join(', ')}` });
    });
  }
  function validateSelloutNoon(rows, errors, warnings) {
    rows.forEach((r, i) => {
      const excelRow = i + 5;
      const metric = r['Metric'];
      if (isBlank(metric)) return;
      if (!isNumeric(r['This Month'])) errors.push({ sheet: 'Sellout Noon', row: excelRow, column: 'This Month', message: `"${r['This Month']}" is not a valid number for "${metric}"` });
    });
  }

  function runValidation(sheetJsonFn) {
    const errors = [], warnings = [];
    const target = sheetJsonFn('Target vs Achievement').filter(r => !isBlank(r['SKU Name']) || Object.values(r).some(v => !isBlank(v)));
    const dist = sheetJsonFn('Distribution Plan').filter(r => Object.values(r).some(v => !isBlank(v)));
    const stock = sheetJsonFn('Stock Status').filter(r => Object.values(r).some(v => !isBlank(v)));
    const sellout = sheetJsonFn('Sellout Noon').filter(r => Object.values(r).some(v => !isBlank(v)));
    if (target.length) validateTargetVsAchievement(target, errors, warnings);
    if (dist.length) validateDistributionPlan(dist, errors, warnings);
    if (stock.length) validateStockStatus(stock, errors, warnings);
    if (sellout.length) validateSelloutNoon(sellout, errors, warnings);
    if (!target.length && !dist.length && !stock.length && !sellout.length) {
      errors.push({ sheet: '(file)', row: null, column: null, message: 'No recognized tabs found — check tab names match the template exactly' });
    }
    return { errors, warnings };
  }

  // ============================================================
  // BRANDS — list / create / read / update
  // ============================================================

  // List brands visible to the caller (all for admin, granted-only otherwise)
  app.get('/api/brands', requireAuth, async function (req, res) {
    try {
      let rows;
      const BRAND_COLS = `b.id, b.name, b.slug, b.accent_color, b.logo_url, (b.logo_data IS NOT NULL) AS has_logo, b.tagline, b.status, b.created_by, b.created_at, b.updated_at`;
      if (isAdmin(req.user)) {
        const r = await pool.query(`
          SELECT ${BRAND_COLS},
            (SELECT string_agg(DISTINCT COALESCE(u.full_name, u.username), ', ')
             FROM brand_user_access a JOIN users u ON u.id = a.user_id
             WHERE a.brand_id = b.id AND a.role = 'manager') AS managers,
            (SELECT MAX(s.updated_at) FROM brand_sections s WHERE s.brand_id = b.id) AS content_updated_at
          FROM brands b ORDER BY b.name`);
        rows = r.rows;
      } else {
        const r = await pool.query(
          `SELECT ${BRAND_COLS}, a.role AS my_role FROM brands b
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
      const brandR = await pool.query(
        `SELECT id, name, slug, accent_color, logo_url, (logo_data IS NOT NULL) AS has_logo, tagline, status, created_by, created_at, updated_at FROM brands WHERE id=$1`,
        [req.params.id]
      );
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
  app.put('/api/brands/:id', requireAuth, requireRole('superadmin'), async function (req, res) {
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
         WHERE id=$6
         RETURNING id, name, slug, accent_color, logo_url, (logo_data IS NOT NULL) AS has_logo, tagline, status, created_by, created_at, updated_at`,
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
  // EXCEL TEMPLATE — live download (Phase 1 item 2)
  // ============================================================
  app.get('/api/brands/:id/excel-template', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const ExcelJS = require('exceljs');
      const wb = new ExcelJS.Workbook();
      const FONT = { name: 'Arial' };
      const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A76' } };
      const headerFont = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      const inputFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      const exampleFont = { name: 'Arial', italic: true, size: 10, color: { argb: 'FF999999' } };
      const titleFont = { name: 'Arial', bold: true, size: 14, color: { argb: 'FF1F7A76' } };
      const legendFont = { name: 'Arial', italic: true, size: 9, color: { argb: 'FF666666' } };

      function addSheet(name, title, headers, colWidths, exampleRow, extraRows) {
        const ws = wb.addWorksheet(name);
        ws.mergeCells(1, 1, 1, headers.length);
        ws.getCell(1, 1).value = title;
        ws.getCell(1, 1).font = titleFont;
        ws.mergeCells(2, 1, 2, headers.length);
        ws.getCell(2, 1).value = 'Yellow cells = fill in your data. Do not rename this tab or move columns.';
        ws.getCell(2, 1).font = legendFont;
        headers.forEach((h, i) => {
          const c = ws.getCell(4, i + 1);
          c.value = h; c.fill = headerFill; c.font = headerFont; c.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        if (exampleRow) exampleRow.forEach((v, i) => { const c = ws.getCell(5, i + 1); c.value = v; c.font = exampleFont; });
        for (let r = 6; r <= 6 + (extraRows || 9); r++) {
          for (let c = 1; c <= headers.length; c++) { const cell = ws.getCell(r, c); cell.fill = inputFill; }
        }
        colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
        return ws;
      }

      addSheet('Target vs Achievement', 'Sales Performance — IMS: Target vs Achievement',
        ['SKU Name', 'Target (Units)', 'Achieved (Units)', 'Period (e.g. Jul-2026)'],
        [26, 16, 18, 20], ['Baby Wipes 80s', 12000, 11280, 'Jul-2026']);
      addSheet('Distribution Plan', 'Sales Performance — IMS: Distribution Plan vs Actuals',
        ['Region / Channel', 'Planned Stores', 'Actual Stores Reached'],
        [26, 18, 20], ['Dubai HoReCa', 120, 108]);
      addSheet('Top Selling SKUs', 'Sales Performance — IMS: Top Selling SKUs',
        ['Rank', 'SKU Name', 'Units Sold'], [10, 26, 16], [1, 'Diaper Pants M', 8420]);
      const selloutWs = addSheet('Sellout Noon', 'Sales Performance — Sellout (Noon)',
        ['Metric', 'This Month', 'Last Month', 'Unit'], [22, 14, 14, 10], null, 0);
      [['Sellout Units', 6140, 5200, 'units'], ['Sellout Revenue', 214000, 187000, 'AED'],
       ['Conversion Rate', 3.8, 3.4, '%'], ['Stock Cover', 21, 18, 'days']].forEach((row, idx) => {
        row.forEach((v, i) => { const c = selloutWs.getCell(5 + idx, i + 1); c.value = v; c.fill = inputFill; });
      });
      addSheet('Stock Status', 'SKU-wise Stock Status — Value, Volume, Ageing, Shelf Life',
        ['SKU Name', 'Stock Value (AED)', 'Volume (Units)', 'Ageing Bucket (0-30/31-60/61+)', 'Shelf Life Status (Healthy/Monitor/Near Expiry)', 'Notes'],
        [24, 18, 16, 30, 34, 30], ['Baby Lotion 200ml', 29750, 3960, '61+', 'Near Expiry', 'Push before month end']);

      const readme = wb.addWorksheet('READ ME');
      readme.getCell(1, 1).value = 'Brand Data Upload Template';
      readme.getCell(1, 1).font = titleFont;
      const instructions = [
        '1. Fill in the yellow cells on each tab below with this period\'s numbers.',
        '2. Do not rename tabs, reorder columns, or delete headers.',
        '3. Upload this file in the brand\'s presentation under "Upload & refresh".',
        '4. Every upload is checked for errors before anything goes live — you will see a preview first.',
        '5. Narrative sections (SWOT, Initiatives, Documents) are edited directly in the app, not here.'
      ];
      instructions.forEach((line, i) => { readme.getCell(4 + i, 1).value = line; readme.getCell(4 + i, 1).font = { name: 'Arial', size: 10.5 }; });
      readme.getColumn(1).width = 95;

      const buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="Brand_Data_Template.xlsx"');
      res.send(Buffer.from(buf));
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // EXCEL UPLOAD — creates a DRAFT only. Never touches the live presentation.
  // (Phase 1 items 3, 4, 5, 10)
  // ============================================================
  // Expected tabs match the template above:
  //   "Target vs Achievement"   -> kpi_hero + sku_rings
  //   "Distribution Plan"       -> folds into kpi_hero
  //   "Stock Status"            -> stock_table
  //   "Sellout Noon"            -> kpi_hero (its own slot)
  // Column headers start at spreadsheet row 4 (range:3).

  app.post('/api/brands/:id/upload-excel', requireAuth, requireBrandManage(), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      const { presentation_name, reporting_month, reporting_year } = req.body;
      if (!presentation_name) return res.status(400).json({ error: 'Presentation name is required' });
      if (!reporting_month || !reporting_year) return res.status(400).json({ error: 'Reporting month and year are required' });

      const brandId = req.params.id;
      const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
      const sheetJson = (name) => {
        if (!wb.Sheets[name]) return [];
        return XLSX.utils.sheet_to_json(wb.Sheets[name], { range: 3, defval: '' });
      };

      const { errors, warnings } = runValidation(sheetJson);

      // Build the excel-driven sections purely in memory — nothing written to
      // brand_sections yet. If there are blocking errors we still return a
      // preview-able draft (so the manager can SEE what's wrong), but the
      // publish endpoint refuses to go live while errors exist.
      const excelSections = [];
      const targetRows = sheetJson('Target vs Achievement').filter(r => !isBlank(r['SKU Name']));
      if (targetRows.length) {
        const items = targetRows.map(r => ({ sku: r['SKU Name'], target: Number(r['Target (Units)']) || 0, achieved: Number(r['Achieved (Units)']) || 0 }));
        excelSections.push({ template_type: 'sku_rings', section_order: 2, chapter_tag: '01 — Sales Performance · IMS', title: 'Top SKU performance', data: { items } });
        const totalTarget = items.reduce((a, b) => a + b.target, 0);
        const totalAchieved = items.reduce((a, b) => a + b.achieved, 0);
        const kpis = [{ label: 'Achieved / Target', value: totalAchieved, compare_value: totalTarget, unit: 'units' }];
        const distRows = sheetJson('Distribution Plan').filter(r => !isBlank(r['Region / Channel']));
        if (distRows.length) {
          const planned = distRows.reduce((a, r) => a + (Number(r['Planned Stores']) || 0), 0);
          const actual = distRows.reduce((a, r) => a + (Number(r['Actual Stores Reached']) || 0), 0);
          kpis.push({ label: 'Store Coverage', value: actual, compare_value: planned, unit: 'stores' });
        }
        excelSections.push({ template_type: 'kpi_hero', section_order: 1, chapter_tag: '01 — Sales Performance · IMS', title: 'Target vs. Achievement', data: { kpis } });
      }
      const stockRows = sheetJson('Stock Status').filter(r => !isBlank(r['SKU Name']));
      if (stockRows.length) {
        const rows = stockRows.map(r => ({
          sku: r['SKU Name'], value: Number(r['Stock Value (AED)']) || 0, volume: Number(r['Volume (Units)']) || 0,
          ageing: r['Ageing Bucket (0-30/31-60/61+)'] || '0-30', shelf_status: r['Shelf Life Status (Healthy/Monitor/Near Expiry)'] || 'Healthy'
        }));
        excelSections.push({ template_type: 'stock_table', section_order: 3, chapter_tag: '03 — SKU-wise Stock Status', title: 'Value, ageing & shelf life', data: { rows } });
      }
      const selloutRows = sheetJson('Sellout Noon').filter(r => !isBlank(r['Metric']));
      if (selloutRows.length) {
        const find = (m) => { const row = selloutRows.find(r => String(r['Metric']).toLowerCase().includes(m)); return row && isNumeric(row['This Month']) ? Number(row['This Month']) : null; };
        const kpis = [];
        const u = find('sellout units'); if (u !== null) kpis.push({ label: 'Sellout Units MTD', value: u, unit: 'units' });
        const rv = find('sellout revenue'); if (rv !== null) kpis.push({ label: 'Sellout Revenue', value: rv, unit: 'AED' });
        const cv = find('conversion'); if (cv !== null) kpis.push({ label: 'Conversion Rate', value: cv, unit: '%' });
        const sc = find('stock cover'); if (sc !== null) kpis.push({ label: 'Stock Cover', value: sc, unit: 'days' });
        if (kpis.length) excelSections.push({ template_type: 'kpi_hero', section_order: 5, chapter_tag: '02 — Sales Performance · Sellout (Noon)', title: 'Noon.com channel', data: { kpis } });
      }

      if (excelSections.length === 0 && errors.length === 0) {
        errors.push({ sheet: '(file)', row: null, column: null, message: 'No recognized data found in any tab' });
      }

      // Merge with the manager's non-Excel sections (cover, SWOT, initiatives,
      // documents) so the draft represents the FULL deck, not just the refreshed part.
      const liveSections = await pool.query(
        `SELECT template_type, section_order, chapter_tag, title, data FROM brand_sections WHERE brand_id=$1 AND section_order NOT IN (1,2,3,5)`,
        [brandId]
      );
      const fullSectionsData = [...liveSections.rows, ...excelSections].sort((a, b) => a.section_order - b.section_order);

      // Only one active (unpublished) draft per brand at a time — a fresh
      // upload supersedes whatever draft existed before, so nothing lingers.
      await pool.query(`UPDATE brand_versions SET status='cancelled' WHERE brand_id=$1 AND status='draft'`, [brandId]);

      const verNumResult = await pool.query(`SELECT COALESCE(MAX(version_number),0)+1 AS n FROM brand_versions WHERE brand_id=$1`, [brandId]);
      const versionNumber = verNumResult.rows[0].n;

      const draft = await pool.query(
        `INSERT INTO brand_versions (brand_id, version_number, presentation_name, reporting_month, reporting_year, status, sections_data, validation, source_file_name, created_by)
         VALUES ($1,$2,$3,$4,$5,'draft',$6,$7,$8,$9) RETURNING id, version_number, presentation_name, reporting_month, reporting_year, status, validation, created_at`,
        [brandId, versionNumber, presentation_name, reporting_month, reporting_year, JSON.stringify(fullSectionsData), JSON.stringify({ errors, warnings }), req.file.originalname, req.user.uid]
      );

      await pool.query(
        `INSERT INTO brand_excel_uploads (brand_id, file_name, uploaded_by, status, sections_updated, error_message)
         VALUES ($1,$2,$3,$4,$5,$6)`,
        [brandId, req.file.originalname, req.user.uid, errors.length ? 'partial' : 'success', excelSections.map(s => s.template_type),
         errors.length ? `${errors.length} error(s) found — draft created but not published` : null]
      );
      await auditLog(req.user.uid, req.user.username, 'brand_draft_created', `brand ${brandId}: v${versionNumber} "${presentation_name}" (${errors.length} errors, ${warnings.length} warnings)`, req.ip);

      res.json({ draft: draft.rows[0], sections_preview: fullSectionsData });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // VERSIONS — list, detail, publish, cancel (Phase 1 items 5, 6, 7, 8, 10)
  // ============================================================

  app.get('/api/brands/:id/versions', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const r = await pool.query(
        `SELECT v.id, v.version_number, v.presentation_name, v.reporting_month, v.reporting_year, v.status,
                v.validation, v.source_file_name, v.created_at, v.published_at,
                cu.full_name AS created_by_name, pu.full_name AS published_by_name
         FROM brand_versions v
         LEFT JOIN users cu ON cu.id = v.created_by
         LEFT JOIN users pu ON pu.id = v.published_by
         WHERE v.brand_id = $1 AND v.status != 'cancelled' ORDER BY v.version_number DESC`,
        [req.params.id]
      );
      res.json({ versions: r.rows });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.get('/api/brands/:id/versions/:versionId', requireAuth, requireBrandView(), async function (req, res) {
    try {
      const r = await pool.query(`SELECT * FROM brand_versions WHERE id=$1 AND brand_id=$2`, [req.params.versionId, req.params.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Version not found' });
      res.json({ version: r.rows[0] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Confirm & Publish — refuses if the draft has any blocking errors.
  // Supersedes (never deletes) the previously published version.
  app.post('/api/brands/:id/versions/:versionId/publish', requireAuth, requireBrandManage(), async function (req, res) {
    const client = await pool.connect();
    try {
      const brandId = req.params.id;
      const vr = await client.query(`SELECT * FROM brand_versions WHERE id=$1 AND brand_id=$2`, [req.params.versionId, brandId]);
      const version = vr.rows[0];
      if (!version) return res.status(404).json({ error: 'Version not found' });
      if (version.status !== 'draft') return res.status(400).json({ error: `This version is already ${version.status}` });
      if ((version.validation.errors || []).length > 0) {
        return res.status(400).json({ error: `Cannot publish — ${version.validation.errors.length} unresolved error(s). Fix the file and re-upload.` });
      }

      await client.query('BEGIN');

      // Supersede whatever was published before.
      await client.query(`UPDATE brand_versions SET status='superseded' WHERE brand_id=$1 AND status='published'`, [brandId]);

      // Apply this version's sections as the live presentation, snapshotting
      // each replaced section into brand_section_history first.
      for (const sec of version.sections_data) {
        const existing = await client.query(
          `SELECT id, data, data_source FROM brand_sections WHERE brand_id=$1 AND section_order=$2`,
          [brandId, sec.section_order]
        );
        if (existing.rows[0]) {
          await client.query(
            `INSERT INTO brand_section_history (section_id, data, data_source, replaced_by) VALUES ($1,$2,$3,$4)`,
            [existing.rows[0].id, existing.rows[0].data, existing.rows[0].data_source, req.user.uid]
          );
          await client.query(
            `UPDATE brand_sections SET template_type=$1, chapter_tag=$2, title=$3, data=$4, data_source='excel_upload', last_refreshed_at=NOW(), updated_by=$5, updated_at=NOW() WHERE id=$6`,
            [sec.template_type, sec.chapter_tag, sec.title, JSON.stringify(sec.data), req.user.uid, existing.rows[0].id]
          );
        } else {
          await client.query(
            `INSERT INTO brand_sections (brand_id, template_type, section_order, chapter_tag, title, data, data_source, last_refreshed_at, updated_by)
             VALUES ($1,$2,$3,$4,$5,$6,'excel_upload',NOW(),$7)`,
            [brandId, sec.template_type, sec.section_order, sec.chapter_tag, sec.title, JSON.stringify(sec.data), req.user.uid]
          );
        }
      }

      await client.query(
        `UPDATE brand_versions SET status='published', published_by=$1, published_at=NOW(), notes=$2 WHERE id=$3`,
        [req.user.uid, req.body.notes || null, version.id]
      );

      await client.query('COMMIT');
      await auditLog(req.user.uid, req.user.username, 'brand_version_published', `brand ${brandId}: v${version.version_number} "${version.presentation_name}"`, req.ip);
      res.json({ ok: true, version_id: version.id, version_number: version.version_number });
    } catch (e) {
      await client.query('ROLLBACK');
      res.status(500).json({ error: e.message });
    } finally {
      client.release();
    }
  });

  // Cancel a draft (Phase 1 item 7 — the "Cancel" option)
  app.delete('/api/brands/:id/versions/:versionId', requireAuth, requireBrandManage(), async function (req, res) {
    try {
      const r = await pool.query(`SELECT status FROM brand_versions WHERE id=$1 AND brand_id=$2`, [req.params.versionId, req.params.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Version not found' });
      if (r.rows[0].status !== 'draft') return res.status(400).json({ error: 'Only a draft can be cancelled' });
      await pool.query(`UPDATE brand_versions SET status='cancelled' WHERE id=$1`, [req.params.versionId]);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // ============================================================
  // SNAPSHOTS — legacy dated archive (superseded by brand_versions above,
  // kept so nothing that already depends on it breaks)
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

  // ============================================================
  // LOGO — upload through the UI (Phase 1 item 9), stored as BYTEA like documents
  // ============================================================
  app.post('/api/brands/:id/logo', requireAuth, requireBrandManage(), upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      if (!req.file.mimetype.startsWith('image/')) return res.status(400).json({ error: 'Logo must be an image file' });
      await pool.query(
        `UPDATE brands SET logo_data=$1, logo_mime=$2, updated_at=NOW() WHERE id=$3`,
        [req.file.buffer, req.file.mimetype, req.params.id]
      );
      await auditLog(req.user.uid, req.user.username, 'brand_logo_updated', `brand ${req.params.id}`, req.ip);
      res.json({ ok: true, logo_url: `/api/brands/${req.params.id}/logo` });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Public-within-app logo serving — <img src> can't send auth headers, so this
  // checks the same session token passed as a query param, same trust boundary
  // as everywhere else, just a different transport for this one image tag.
  app.get('/api/brands/:id/logo', async function (req, res) {
    try {
      const token = req.headers['x-auth-token'] || req.query.token;
      if (!token) return res.status(401).end();
      const sess = await pool.query(`SELECT user_id FROM sessions WHERE token=$1 AND expires_at>NOW()`, [token]);
      if (!sess.rows[0]) return res.status(401).end();
      const r = await pool.query(`SELECT logo_data, logo_mime FROM brands WHERE id=$1`, [req.params.id]);
      if (!r.rows[0] || !r.rows[0].logo_data) return res.status(404).end();
      res.setHeader('Content-Type', r.rows[0].logo_mime || 'image/png');
      res.setHeader('Cache-Control', 'private, max-age=300');
      res.send(r.rows[0].logo_data);
    } catch (e) {
      res.status(500).end();
    }
  });

  // ============================================================
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

  app.get('/api/brands/:id/access', requireAuth, requireRole('superadmin'), async function (req, res) {
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

  app.post('/api/brands/:id/access', requireAuth, requireRole('superadmin'), async function (req, res) {
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

  app.delete('/api/brands/:id/access/:userId', requireAuth, requireRole('superadmin'), async function (req, res) {
    try {
      await pool.query(`DELETE FROM brand_user_access WHERE brand_id=$1 AND user_id=$2`, [req.params.id, req.params.userId]);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  console.log('Brand module: routes mounted');
};
