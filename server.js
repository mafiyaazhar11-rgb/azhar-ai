const express = require('express');
const multer = require('multer');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 150 * 1024 * 1024 } });

// Security headers (HSTS, X-Frame-Options, X-Content-Type-Options, etc.)
app.use(helmet({ contentSecurityPolicy: false })); // CSP off since inline scripts are used across the existing dashboards

// Restrict cross-origin API access to known app domains only
const ALLOWED_ORIGINS = [
  'https://azr-operations.com',
  'https://azhar-ai-la1l.onrender.com',
  'http://localhost:3000'
];
app.use(cors({
  origin: function (origin, callback) {
    if (!origin || ALLOWED_ORIGINS.includes(origin)) return callback(null, true);
    return callback(new Error('Not allowed by CORS'));
  },
  credentials: true
}));

// Brute-force protection on login: 10 attempts per 15 minutes per IP
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  message: { error: 'Too many login attempts. Please try again in 15 minutes.' },
  standardHeaders: true,
  legacyHeaders: false
});

app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ extended: true, limit: '150mb' }));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

//  PostgreSQL 
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

//  Fallback file storage (if DB not available) 
const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DISPATCH_FILE  = path.join(DATA_DIR, 'dispatch.json');
const REJECTION_FILE = path.join(DATA_DIR, 'rejection.json');

function saveJSON(fp, data) {
  try { fs.writeFileSync(fp, JSON.stringify(data)); } catch(e) { console.error('File save error:', e.message); }
}
function loadJSON(fp) {
  try { if (fs.existsSync(fp)) return JSON.parse(fs.readFileSync(fp, 'utf8')); } catch(e) {}
  return null;
}

//  Init DB tables 
async function initDB() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS dispatch_data (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL UNIQUE,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT DEFAULT 'Admin',
      summary JSONB NOT NULL,
      csv_text TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Transport-team-reported drop counts, uploaded from their own raw drop
    // file (Distinct DROP ID, by date + Bulk/Multi classification). Compared
    // against the app's own dispatch_data.summary.total_drops for the same
    // dates to catch under/over-counting on either side before the monthly
    // invoice arrives.
    //
    // reported_drops = ALL dispatched drops for that date/class, regardless of
    // the transport team's internal TASK STATUS (Completed/Ongoing/Waiting) —
    // a drop that was dispatched already incurred the trip cost whether or not
    // it's been marked done in their system yet. completed_drops is kept only
    // as supplementary context (how much of that day has been closed out),
    // never used to filter the primary count.
    await pool.query(`CREATE TABLE IF NOT EXISTS transport_drop_reconciliation (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL,
      drop_class TEXT NOT NULL,
      reported_drops INT NOT NULL DEFAULT 0,
      completed_drops INT NOT NULL DEFAULT 0,
      task_status_rule TEXT,
      upload_batch_id TEXT,
      uploaded_by TEXT,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      UNIQUE(date_key, drop_class)
    )`);
    await pool.query(`ALTER TABLE transport_drop_reconciliation ADD COLUMN IF NOT EXISTS completed_drops INT NOT NULL DEFAULT 0`);
    // Individual order identifiers from the transport team's file (TASK ID,
    // suffix stripped), per date — lets us diff against this app's own
    // order_tracking table (already populated from daily dispatch uploads)
    // to find the SPECIFIC orders that don't match on either side, not just
    // a day-level count difference.
    await pool.query(`CREATE TABLE IF NOT EXISTS transport_order_ids (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL,
      order_id TEXT NOT NULL,
      upload_batch_id TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      UNIQUE(date_key, order_id)
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS rejection_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT DEFAULT 'Admin',
      file_name TEXT,
      total_orders INT,
      orgs JSONB,
      months JSONB,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS backlog_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS returns_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    // Tracks every order code seen per dispatch date, so the same order code appearing
    // again on a LATER date can be detected as a re-delivery (failed first attempt,
    // re-attempted later) — not just a same-day duplicate.
    await pool.query(`CREATE TABLE IF NOT EXISTS order_tracking (
      id SERIAL PRIMARY KEY,
      order_code TEXT NOT NULL,
      date_key DATE NOT NULL,
      customer TEXT,
      value NUMERIC DEFAULT 0,
      route TEXT,
      org TEXT,
      drop_type TEXT,
      temperature TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS temperature TEXT`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_order_tracking_code ON order_tracking(order_code)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_order_tracking_date ON order_tracking(date_key)`);
    // AUTH TABLES
    await pool.query(`CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      full_name TEXT,
      role TEXT NOT NULL DEFAULT 'viewer',
      dashboards JSONB DEFAULT '["dispatch","rejection","summary","email","invoice","backlog","returns","sales","automation"]'::jsonb,
      active BOOLEAN DEFAULT true,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      created_by TEXT DEFAULT 'system',
      last_login TIMESTAMPTZ,
      must_change_password BOOLEAN DEFAULT false
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS sessions (
      id SERIAL PRIMARY KEY,
      token TEXT UNIQUE NOT NULL,
      user_id INT REFERENCES users(id) ON DELETE CASCADE,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      expires_at TIMESTAMPTZ NOT NULL,
      ip_address TEXT
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS audit_log (
      id SERIAL PRIMARY KEY,
      user_id INT,
      username TEXT,
      action TEXT NOT NULL,
      details TEXT,
      ip_address TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Create or update default super admin
    var adminCheck = await pool.query("SELECT id FROM users WHERE username = 'azhar'");
    var hash = await bcrypt.hash('YAmaha100@', 10);
    if (adminCheck.rows.length === 0) {
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role) VALUES ($1,$2,$3,$4)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      console.log('Default super admin created: azhar / YAmaha100@');
    } else {
      // Always sync password with code on server start
      await pool.query("UPDATE users SET password_hash=$1, active=true WHERE username='azhar'", [hash]);
      console.log('Super admin password synced: azhar / YAmaha100@');
    }
    console.log('DB tables ready');
  } catch(e) {
    console.error('DB init error:', e.message);
  }
}
initDB();

//  DB helpers 
async function dbSaveDispatch(dateKey, uploadedBy, summary, csvText) {
  try {
    await pool.query(`
      INSERT INTO dispatch_data (date_key, uploaded_by, summary, csv_text)
      VALUES ($1, $2, $3, $4)
      ON CONFLICT (date_key) DO UPDATE
      SET uploaded_by=$2, summary=$3, csv_text=$4, uploaded_at=NOW()
    `, [dateKey, uploadedBy, JSON.stringify(summary), csvText?.substring(0, 500000)]);
    return true;
  } catch(e) {
    console.error('DB save dispatch error:', e.message);
    return false;
  }
}

async function dbLoadDispatch() {
  try {
    const res = await pool.query(`SELECT date_key::text, uploaded_at, uploaded_by, summary, csv_text FROM dispatch_data ORDER BY date_key DESC LIMIT 60`);
    return res.rows;
  } catch(e) {
    console.error('DB load dispatch error:', e.message);
    return [];
  }
}

async function dbSaveRejection(uploadedBy, fileName, totalOrders, orgs, months) {
  try {
    await pool.query(`DELETE FROM rejection_data`);
    await pool.query(`
      INSERT INTO rejection_data (uploaded_by, file_name, total_orders, orgs, months)
      VALUES ($1, $2, $3, $4, $5)
    `, [uploadedBy, fileName, totalOrders, orgs, months]);
    return true;
  } catch(e) {
    console.error('DB save rejection error:', e.message);
    return false;
  }
}

async function dbLoadRejection() {
  try {
    const res = await pool.query(`SELECT * FROM rejection_data ORDER BY created_at DESC LIMIT 1`);
    return res.rows[0] || null;
  } catch(e) {
    console.error('DB load rejection error:', e.message);
    return null;
  }
}


var backlogData = null;
var BACKLOG_FILE = path.join(DATA_DIR, 'backlog.json');

async function dbSaveBacklog(uploadedBy, fileName, totalOrders, summary) {
  try {
    await pool.query('DELETE FROM backlog_data');
    await pool.query(
      'INSERT INTO backlog_data (uploaded_by, file_name, total_orders, summary) VALUES ($1, $2, $3, $4)',
      [uploadedBy, fileName, totalOrders, JSON.stringify(summary)]
    );
    return true;
  } catch(e) {
    console.error('DB save backlog error:', e.message);
    return false;
  }
}

async function loadBacklogFromDB() {
  try {
    var res = await pool.query('SELECT * FROM backlog_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      backlogData = { uploadedAt: res.rows[0].uploaded_at, uploadedBy: res.rows[0].uploaded_by, fileName: res.rows[0].file_name, totalOrders: res.rows[0].total_orders, summary: res.rows[0].summary };
      console.log('Loaded backlog from DB');
      return true;
    }
  } catch(e) { console.error('DB load backlog:', e.message); }
  var saved = loadJSON(BACKLOG_FILE);
  if (saved) { backlogData = saved; console.log('Loaded backlog from file'); }
  return false;
}
loadBacklogFromDB();

app.post('/api/backlog/upload', upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    var summary = (typeof req.body.summary === 'object') ? req.body.summary : JSON.parse(req.body.summary || '{}');
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var fileName = req.file.originalname || 'backlog.xlsx';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    backlogData = { uploadedAt: new Date().toISOString(), uploadedBy: uploadedBy, fileName: fileName, totalOrders: totalOrders, summary: summary };
    var dbOk = await dbSaveBacklog(uploadedBy, fileName, totalOrders, summary);
    saveJSON(BACKLOG_FILE, backlogData);
    console.log('Backlog saved:', totalOrders, 'orders', dbOk ? '(DB+file)' : '(file only)');
    res.json({ success: true });
  } catch(e) {
    console.error('Backlog upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.delete('/api/backlog/clear', async function(req, res) {
  try {
    await pool.query('DELETE FROM backlog_data');
    backlogData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var BACKLOG_SUMMARY_VERSION = 'v3'; // Increment when shortCat logic changes
app.get('/api/backlog/status', function(req, res) {
  if (!backlogData) return res.json({ hasData: false });
  // If summary version doesn't match, force re-upload
  if (!backlogData.summary || backlogData.summary.version !== BACKLOG_SUMMARY_VERSION) {
    return res.json({ hasData: false, reason: 'version_mismatch' });
  }
  res.json({ hasData: true, uploadedAt: backlogData.uploadedAt, uploadedBy: backlogData.uploadedBy, fileName: backlogData.fileName, totalOrders: backlogData.totalOrders, summary: backlogData.summary });
});


app.get('/health', function(req, res) {
  res.json({ status: 'ok', time: new Date().toISOString(), db: !!process.env.DATABASE_URL });
});

//  HELPERS 
function toStr(v) { return String(v == null ? '' : v).trim(); }

// Same DD/MM/YYYY-safe date parser used in the pallet module — JS's bare
// `new Date(string)` assumes MM/DD/YYYY and silently mis-parses UAE-format
// text dates like "14/07/2026". Needed here too since transport-team files
// mix real Excel date cells with text-formatted ones.
function toDateStrGeneric(val) {
  if (val === undefined || val === null || val === '') return null;
  if (val instanceof Date) { return isNaN(val.getTime()) ? null : val.toISOString().slice(0, 10); }
  if (typeof val === 'number') {
    var d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  var str = String(val).trim();
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

function normaliseType(raw) {
  var t = toStr(raw).toUpperCase().replace(/\s+/g, ' ').trim();
  if (t === 'FOOD' || t.startsWith('FOOD')) return 'food';
  if (t.includes('NON-FOOD') || t.includes('NON FOOD')) return 'nonfood';
  if (t === 'GSEB' || t === 'SHARKNINJA') return 'nonfood';
  if (t === '3PL' || t === '3 PL' || t === 'HCP') return '3pl';
  if (t === 'VAN') return 'van';
  return t.toLowerCase();
}

function normaliseCity(raw) {
  var c = toStr(raw).toLowerCase();
  if (c.includes('abu dhabi')) return 'Abu Dhabi';
  if (c.includes('hatta')) return 'Hatta';
  if (c.includes('dubai')) return 'Dubai';
  if (c.includes('sharjah')) return 'Sharjah';
  if (c.includes('ajman')) return 'Ajman';
  if (c.includes('fujairah')) return 'Fujairah';
  if (c.includes('al ain') || c.includes('al-ain')) return 'Al Ain';
  if (c.includes('ras al') || c === 'rak') return 'Ras Al Khaimah';
  if (c.includes('umm')) return 'Umm Al Quwain';
  var s = toStr(raw);
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

// The CITY column is sometimes wrong (e.g. an internal transfer address that literally
// contains "Sharjah" in the text gets logged under CITY="Dubai"). Scan the full address
// text for a city name and prefer that over a mismatched CITY column value.
function detectCityFromAddress(addressText, cityColumnValue) {
  var fallback = normaliseCity(cityColumnValue);
  var addr = toStr(addressText).toLowerCase();
  if (!addr) return fallback;
  var found = null;
  if (addr.includes('abu dhabi')) found = 'Abu Dhabi';
  else if (addr.includes('hatta')) found = 'Hatta';
  else if (addr.includes('sharjah')) found = 'Sharjah';
  else if (addr.includes('ajman')) found = 'Ajman';
  else if (addr.includes('fujairah')) found = 'Fujairah';
  else if (addr.includes('al ain') || addr.includes('al-ain')) found = 'Al Ain';
  else if (addr.includes('ras al khaimah') || addr.includes('rak,') || addr.endsWith('rak')) found = 'Ras Al Khaimah';
  else if (addr.includes('umm al quwain')) found = 'Umm Al Quwain';
  else if (addr.includes('dubai')) found = 'Dubai';
  // If the address text clearly names a different city than the CITY column, trust the address text.
  return found || fallback;
}

// Transport team's FY26 rate card (AED per vehicle, per day/trip) — provided directly by
// transport, not estimated. Matches on distinctive keywords so variations in how the truck-type
// column gets typed ("Ambient-Multi", "AMBIENT MULTI", etc.) still resolve to the right rate.
var TRUCK_RATE_CARD = [
  { keywords: ['FROZEN', 'MULTI'], label: 'Frozen - Multi', rate: 120 },
  { keywords: ['FROZEN', '4 TON'], label: 'Frozen - Bulk 4 Ton', rate: 850 },
  { keywords: ['FROZEN', '10 TON'], label: 'Frozen - Bulk 10 Ton', rate: 1350 },
  { keywords: ['AMBIENT', 'MULTI'], label: 'Ambient - Multi', rate: 104 },
  { keywords: ['AMBIENT', '4 TON'], label: 'Ambient - Bulk 4 Ton', rate: 750 },
  { keywords: ['AMBIENT', '10 TON'], label: 'Ambient - Bulk 10 Ton', rate: 950 },
  { keywords: ['AMBIENT', '40'], label: 'Ambient - Bulk 40 FT', rate: 1200 },
  { keywords: ['E-COMMERCE'], label: 'E-commerce', rate: 20 },
  { keywords: ['ECOMMERCE'], label: 'E-commerce', rate: 20 },
  { keywords: ['EXCLUSIVE', '1 TON'], label: 'Exclusive 1 Ton', rate: 550 },
  { keywords: ['EXCLUSIVE', '4 TON'], label: 'Exclusive 4 Ton', rate: 750 }
];
function matchTruckRate(rawTruckType) {
  var u = toStr(rawTruckType).toUpperCase();
  if (!u) return null;
  for (var i = 0; i < TRUCK_RATE_CARD.length; i++) {
    var entry = TRUCK_RATE_CARD[i];
    var allMatch = entry.keywords.every(function(kw) { return u.indexOf(kw) !== -1; });
    if (allMatch) return entry;
  }
  return null;
}

// ── Vehicle Master fallback: when a drop has no truck-type text but does have a vehicle
// plate number, look up that vehicle's tonnage + Chiller/Frozen/Ambient from the uploaded
// Vehicle Master, combine with that vehicle's drop-count today (1 drop = Bulk, 2+ = Multi,
// per transport team's rule), and reuse the same rate card via a synthetic keyword string.
function normalizeVehicleNoForLookup(raw) {
  return String(raw || '').toUpperCase().replace(/\s+/g, '');
}
function vehicleTempBucket(vehicleTypeRaw) {
  var u = toStr(vehicleTypeRaw).toUpperCase();
  if (!u) return null;
  if (u.indexOf('FROZEN') !== -1 || u.indexOf('FREEZER') !== -1) return 'FROZEN';
  if (u.indexOf('CHILL') !== -1 || u.indexOf('AMBIENT') !== -1) return 'AMBIENT'; // Chiller priced as Ambient per transport team
  return null; // Dry / Open Pick-up / Car / Bus / Other — no rate applies
}
function vehicleTonnageBucket(vehTypeRaw) {
  var u = toStr(vehTypeRaw).toUpperCase().replace(/\s+/g, '');
  if (u.indexOf('4.2T') !== -1 || u === '4T') return '4 TON';
  if (u.indexOf('10T') !== -1) return '10 TON';
  return null; // 3T / 12T — genuinely no Bulk tier for these on the FY26 rate card
}
function lookupVehicleMasterByAnyId(vehicleId) {
  var norm = normalizeVehicleNoForLookup(vehicleId);
  var v = VEHICLE_MASTER_MAP[norm];
  if (!v) {
    var digitGroups = toStr(vehicleId).match(/\d+/g);
    if (digitGroups && digitGroups.length) {
      v = VEHICLE_MASTER_MAP['DIGITS:' + digitGroups[digitGroups.length - 1]];
    }
  }
  return v || null;
}
function matchRateViaVehicleMaster(vehicleId, dropCountForVehicle) {
  var v = lookupVehicleMasterByAnyId(vehicleId);
  if (!v) return null;
  var tempBucket = vehicleTempBucket(v.vehicle_type_raw);
  if (!tempBucket) return null;
  var isOneTon = toStr(v.veh_type).toUpperCase().replace(/\s+/g, '') === '1T';
  // 1 Ton has no dedicated Multi/Bulk-4-Ton style tier on the FY26 rate card — "Exclusive 1 Ton"
  // (550 AED) is the only rate that exists for this tonnage at all, so it's used regardless of
  // drop count for 1-ton vehicles specifically.
  if (isOneTon) return matchTruckRate('EXCLUSIVE 1 TON');
  var tier = dropCountForVehicle >= 2 ? 'MULTI' : vehicleTonnageBucket(v.veh_type);
  if (!tier) return null;
  return matchTruckRate(tempBucket + ' ' + tier);
}

function extractDriverName(contact) {
  var s = toStr(contact);
  if (!s) return '';
  // Transport staff sometimes already label these explicitly, e.g. "Hired Driver-Ayaz 566" —
  // that's a complete, meaningful identifier already; keep it whole instead of running it
  // through the phone-number-stripping logic below, which would otherwise discard the name
  // and collapse every distinct hired driver into one generic "Hired Driver" entry.
  if (/^hired driver/i.test(s)) return s.trim();
  if (/^[A-Za-z][A-Za-z\s]{2,}$/.test(s)) return s.trim();
  var m = s.match(/^([A-Za-z][A-Za-z\s]{2,29})(?:\s*[-+\d])/);
  if (m) return m[1].trim();
  var parts = s.split(/[-+\d]/);
  var name = (parts[0] || '').trim();
  if (name.length > 2) return name;
  // No name at all in the source data — just a bare phone number/ID. Label it clearly
  // instead of silently displaying the number, so it reads as "name missing", not a bug.
  if (/^\d+$/.test(s)) return 'Hired Driver (ID: ' + s + ')';
  return s.trim();
}

function stripBranch(name) {
  var base = toStr(name);
  var kws = [',Branch', ', Branch', ',Br.', ', Br.', ' -Branch', ',CPD', ' CPD', '- Branch', '-Branch'];
  for (var i = 0; i < kws.length; i++) {
    var idx = base.toLowerCase().indexOf(kws[i].toLowerCase());
    if (idx > 3) { base = base.substring(0, idx).trim(); break; }
  }
  return base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
}

function findDataSheet(wb) {
  var bestSheet = wb.SheetNames[0], bestRows = 0;
  for (var i = 0; i < wb.SheetNames.length; i++) {
    var name = wb.SheetNames[i];
    var ws = wb.Sheets[name];
    if (!ws['!ref']) continue;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var r = range.e.r - range.s.r;
    if (r > bestRows) { bestRows = r; bestSheet = name; }
  }
  console.log('Using sheet:', bestSheet, 'rows:', bestRows);
  return bestSheet;
}

//  DISPATCH PARSER 
function parseDispatch(buffer) {
  var wb = XLSX.read(buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false, cellFormula: false });
  var sheetName = findDataSheet(wb);
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  if (!rows.length) return null;

  function findCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
    }) || null;
  }
  // Exact (not fuzzy-substring) header match, kept as a fallback only. Confirmed directly
  // against real files: "ORDER CODE" (e.g. "4256142DCV130726SO") stays IDENTICAL across
  // days for the same re-delivered order, so it's the primary identifier (see findCol call
  // below). The plain "ORDER" column isn't always present in every day's export, so it's
  // only used when "ORDER CODE" is missing entirely.
  function findExactCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.trim().toUpperCase() === n.toUpperCase(); });
    }) || null;
  }

  var C = {
    route:    findCol('ROUTE'),
    city:     findCol('CITY', 'AREA'),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount:   findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    driver:   findCol('DRIVER CONTACT DETAILS', 'DRIVERS NAME', 'DRIVER NAME', 'DRIVER CONTACT', 'DRIVER_CONTACT', 'DRIVER_ID'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    address:  findCol('CUSTOMER ADDRESS', 'ADDRESS'),
    keep:     findCol('KEEP TOGETHER', 'KEEP_TOGETHER', 'KEEPTOGETHER', 'KEEP'),
    type:     findCol('TYPE'),
    temperature: findCol('TEMPERATURE', 'TEMP'),
    vehicleId: findCol('VEHICLE_ID', 'VEHICLE ID', 'VEHICLE'),
    truckType: findCol('TRUCK TYPE', 'TRUCK_TYPE', 'VEHICLE TYPE', 'VEHICLE_TYPE', 'DROP TYPE', 'DROP_TYPE'),
    orderCode: findCol('ORDER CODE', 'ORDER_CODE') || findExactCol('ORDER', 'ORDER NUMBER', 'ORDER_NUMBER') || findCol('ORDER '),
    org:      findCol('ORG') || findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU')
  };
  console.log('Dispatch cols:', JSON.stringify(C));

  var totalOrders=0, totalValue=0, foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  var frozenOrders=0, frozenValue=0, ambientOrders=0, ambientValue=0;
  var cities={}, customers={}, routes={}, driverSet={};
  var dropsByCity = {}; // one increment per unique (route, location) drop — cannot exceed total_drops by construction
  var dropRecords = {}; // route::loc -> { city, truckType, types:{}, vehicleId } — accumulated across ALL rows for that drop, so region/food-type breakdowns are based on complete data, not just whichever row happened to create the drop first
  var orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, DPS:{o:0,v:0}, DPB:{o:0,v:0}, HCP:{o:0,v:0} };
  var cityTypeCross = {}; // city -> {food, nonfood, pl, van, other}
  var locationVisits = {}; // locationId -> { address, customer, routes:Set(all), ownRoutes:Set(non-3PL) }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    totalOrders++;
    var amt = parseFloat(row[C.amount]) || 0;
    totalValue += amt;
    var type = normaliseType(C.type ? row[C.type] : '');
    var tempForRow = C.temperature ? toStr(row[C.temperature]).toUpperCase() : '';
    var isFrozenRow = tempForRow.indexOf('FROZEN') !== -1;
    if (tempForRow) {
      if (isFrozenRow) { frozenOrders++; frozenValue += amt; }
      else { ambientOrders++; ambientValue += amt; }
    }
    var rawTruckTypeForRow = C.truckType ? toStr(row[C.truckType]) : '';
    if (type === 'food')   { foodOrders++;    foodValue    += amt; }
    else if (type === 'nonfood') { nonFoodOrders++; nonFoodValue += amt; }
    else if (type === '3pl')     { plOrders++; }
    else if (type === 'van')     { vanOrders++; }
    var org = C.org ? toStr(row[C.org]).toUpperCase() : '';
    if (type === '3pl') { orgStats.HCP.o++; orgStats.HCP.v += amt; }
    else if (orgStats[org]) { orgStats[org].o++; orgStats[org].v += amt; }
    else if (org === '3 PL' || org === 'HCP' || org === '3PL') { orgStats.HCP.o++; orgStats.HCP.v += amt; }
    if (C.city && row[C.city]) {
      var city = normaliseCity(row[C.city]);
      if (!cities[city]) cities[city] = { orders:0, value:0 };
      cities[city].orders++; cities[city].value += amt;
      if (!cityTypeCross[city]) cityTypeCross[city] = { food:0, nonfood:0, pl:0, van:0, other:0 };
      if (type === 'food') cityTypeCross[city].food++;
      else if (type === 'nonfood') cityTypeCross[city].nonfood++;
      else if (type === '3pl') cityTypeCross[city].pl++;
      else if (type === 'van') cityTypeCross[city].van++;
      else cityTypeCross[city].other++;
    }
    if (C.customer && row[C.customer]) {
      var cust = toStr(row[C.customer]);
      if (!customers[cust]) customers[cust] = { orders:0, value:0 };
      customers[cust].orders++; customers[cust].value += amt;
    }
    if (C.route && row[C.route]) {
      var route = toStr(row[C.route]);
      if (!routes[route]) routes[route] = { locs:{}, drivers:{}, orders:0, value:0, types:{}, vehicleIds:{} };
      var vehicleIdForRoute = C.vehicleId ? toStr(row[C.vehicleId]) : '';
      if (vehicleIdForRoute) routes[route].vehicleIds[vehicleIdForRoute] = (routes[route].vehicleIds[vehicleIdForRoute] || 0) + 1;
      var routeTypeLabel = (type||'other').charAt(0).toUpperCase()+(type||'other').slice(1);
      if (type === 'nonfood') routeTypeLabel = 'Non-Food';
      else if (type === '3pl') routeTypeLabel = '3PL';
      if (tempForRow) routeTypeLabel += ' (' + (isFrozenRow ? 'Frozen' : 'Ambient') + ')';
      routes[route].types[routeTypeLabel] = (routes[route].types[routeTypeLabel] || 0) + 1;
      var rawLoc = C.location ? toStr(row[C.location]) : '';
      // Internal cash-van transfers all drop at a fixed hub per city (not a real unique customer
      // address per transaction) — collapse them to one drop per city instead of counting every
      // internal order's own LOCATION_ID as a separate physical stop.
      var custForInternal = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
      var isInternalVan = (type === 'van') && custForInternal.indexOf('INTERNAL') !== -1;
      var addrTextForCity = C.address ? toStr(row[C.address]) : '';
      var cityForLoc = detectCityFromAddress(addrTextForCity, C.city ? row[C.city] : '');
      var loc = isInternalVan ? ('INTERNAL-HUB::' + (cityForLoc || 'Unknown')) : rawLoc;
      if (loc) {
        if (!routes[route].locs[loc]) {
          // First time this exact (route, location) pair is seen — this is a genuinely new drop.
          dropsByCity[cityForLoc || 'Unknown'] = (dropsByCity[cityForLoc || 'Unknown'] || 0) + 1;
        }
        routes[route].locs[loc] = 1;
        // Accumulate this drop's full record across every row that contributes to it, so the
        // region + Food/Non-Food breakdown (computed after the loop) sees ALL types on the
        // drop, not just whichever row happened to create it first.
        var dropKey = route + '::' + loc;
        if (!dropRecords[dropKey]) {
          dropRecords[dropKey] = { city: cityForLoc || 'Unknown', truckType: rawTruckTypeForRow, types: {}, vehicleId: '' };
        }
        if (rawTruckTypeForRow && !dropRecords[dropKey].truckType) dropRecords[dropKey].truckType = rawTruckTypeForRow;
        dropRecords[dropKey].types[type || 'other'] = true;
        var vehicleIdValForDrop = C.vehicleId ? toStr(row[C.vehicleId]) : '';
        if (vehicleIdValForDrop) dropRecords[dropKey].vehicleId = vehicleIdValForDrop;
      }
      routes[route].value += amt;
      routes[route].orders++;
      if (C.driver && row[C.driver]) {
        var drvName = extractDriverName(row[C.driver]);
        if (drvName) routes[route].drivers[drvName] = 1;
      }
      // Track which routes visit each physical location, to catch the same address being
      // driven to twice by two different routes on the same day (double-charged drop).
      // Cash/walk-in orders ("**** Cash **** (Dxb)") share a generic placeholder location,
      // not a real fixed address, so they're flagged and excluded from the repeat-visit REPORT
      // further below (but still counted normally for the own-fleet/3PL drop-cost split).
      // Also track the order TYPE per route-visit — Food and Non-Food often can't share a
      // truck, so a "repeat visit" that's actually Food+Non-Food is a legitimate separate
      // trip, not a duplicate/avoidable one. Frozen vs Ambient of the SAME product type also
      // needs a separate truck for temperature control, so that's folded into the same check.
      var truckType = type + (isFrozenRow ? '-frozen' : (tempForRow ? '-ambient' : ''));
      if (loc) {
        var custNameForCash = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
        if (!locationVisits[loc]) {
          locationVisits[loc] = {
            address: C.address ? toStr(row[C.address]) : '',
            customer: C.customer ? toStr(row[C.customer]) : '',
            isCashOrder: custNameForCash.indexOf('CASH') !== -1,
            isInternalHub: loc.indexOf('INTERNAL-HUB::') === 0,
            routes: {}, ownRoutes: {}, typesByRoute: {}, valueByRoute: {}, orderCountByRoute: {}, ordersByRoute: {}
          };
        }
        locationVisits[loc].routes[route] = 1;
        if (type !== '3pl') locationVisits[loc].ownRoutes[route] = 1;
        if (!locationVisits[loc].typesByRoute[route]) locationVisits[loc].typesByRoute[route] = {};
        locationVisits[loc].typesByRoute[route][truckType || 'other'] = true;
        locationVisits[loc].valueByRoute[route] = (locationVisits[loc].valueByRoute[route] || 0) + amt;
        locationVisits[loc].orderCountByRoute[route] = (locationVisits[loc].orderCountByRoute[route] || 0) + 1;
        if (!locationVisits[loc].ordersByRoute[route]) locationVisits[loc].ordersByRoute[route] = [];
        locationVisits[loc].ordersByRoute[route].push({
          order_code: C.orderCode ? toStr(row[C.orderCode]) : '',
          type: type || 'other',
          temperature: isFrozenRow ? 'Frozen' : (tempForRow ? 'Ambient' : ''),
          value: Math.round(amt)
        });
      }
    }
    if (C.driver && row[C.driver]) {
      var keepVal2 = C.keep ? toStr(row[C.keep]) : (C.location ? toStr(row[C.location]) : '');
      var drv = extractDriverName(row[C.driver]) || toStr(row[C.driver]);
      if (drv) driverSet[drv] = 1;
    }
  }

  console.log('TYPE counts food:'+foodOrders+' nonfood:'+nonFoodOrders+' 3pl:'+plOrders+' van:'+vanOrders);

  var byCity = Object.keys(cities).map(function(c) {
    return { city:c, orders:cities[c].orders, value:Math.round(cities[c].value) };
  }).sort(function(a,b) { return b.orders-a.orders; });

  var baseCust = {};
  Object.keys(customers).forEach(function(name) {
    var base = stripBranch(name);
    if (!baseCust[base]) baseCust[base] = { orders:0, value:0 };
    baseCust[base].orders += customers[name].orders;
    baseCust[base].value  += customers[name].value;
  });
  var topCustomers = Object.keys(baseCust).map(function(name) {
    return { name:name, orders:baseCust[name].orders, value:Math.round(baseCust[name].value) };
  }).sort(function(a,b) { return b.value-a.value; }).slice(0,6);

  var topRoutes = Object.keys(routes).map(function(route) {
    var typeMap = routes[route].types || {};
    var topTypesForRoute = Object.keys(typeMap).sort(function(a,b){ return typeMap[b]-typeMap[a]; });
    var hasFood = topTypesForRoute.some(function(t){ return t.indexOf('Food') === 0; });
    var hasNonFood = topTypesForRoute.some(function(t){ return t.indexOf('Non-Food') === 0; });
    var hasFrozen = topTypesForRoute.some(function(t){ return t.indexOf('Frozen') !== -1; });
    var hasAmbient = topTypesForRoute.some(function(t){ return t.indexOf('Ambient') !== -1; });
    var guessedPartition = (hasFood && hasNonFood) || (hasFrozen && hasAmbient);

    // Prefer the actual Vehicle Master partition flag over the route-content guess, whenever
    // we can identify which vehicle ran this route (majority vehicle ID seen on it).
    var isPartitionVehicle = guessedPartition;
    var partitionSource = 'guessed';
    var vids = Object.keys(routes[route].vehicleIds || {});
    if (vids.length) {
      var majorityVid = vids.sort(function(a,b){ return routes[route].vehicleIds[b] - routes[route].vehicleIds[a]; })[0];
      var vm = lookupVehicleMasterByAnyId(majorityVid);
      if (vm && vm.partition_flag) {
        var pf = toStr(vm.partition_flag).trim().toUpperCase();
        if (pf === 'YES' || pf === 'Y') { isPartitionVehicle = true; partitionSource = 'vehicle_master'; }
        else if (pf === 'NO' || pf === 'N') { isPartitionVehicle = false; partitionSource = 'vehicle_master'; }
      }
    }
    return { route:route, orders:routes[route].orders, drops:Object.keys(routes[route].locs).length, driverCount:Object.keys(routes[route].drivers).length, value:Math.round(routes[route].value), types:topTypesForRoute, isPartitionVehicle:isPartitionVehicle, partitionSource:partitionSource };
  }).sort(function(a,b) { return b.drops-a.drops; });

  // Count actual orders per driver (not route drops)
  var driverOrders = {};
  rows.forEach(function(row) {
    var drv = C.driver && row[C.driver] ? extractDriverName(row[C.driver]) : '';
    if (!drv) return;
    var amt = C.amount ? parseFloat(row[C.amount]) || 0 : 0;
    var rawLocId = C.location ? toStr(row[C.location]) : '';
    // Same internal-van hub collapse + address-based city correction used for Route Summary,
    // so the driver leaderboard's drop count matches what's shown everywhere else.
    var typeForDrv = C.type ? normaliseType(row[C.type]) : '';
    var custForDrv = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
    var isInternalVanDrv = (typeForDrv === 'van') && custForDrv.indexOf('INTERNAL') !== -1;
    var cityForDrv = detectCityFromAddress(C.address ? toStr(row[C.address]) : '', C.city ? row[C.city] : '');
    var locId = isInternalVanDrv ? ('INTERNAL-HUB::' + (cityForDrv || 'Unknown')) : rawLocId;
    var custNameForDrv = C.customer ? toStr(row[C.customer]) : '';
    if (!driverOrders[drv]) driverOrders[drv] = {orders:0, drops:{}, value:0, customers:{}, types:{}};
    driverOrders[drv].orders++;
    driverOrders[drv].value += amt;
    if (locId) driverOrders[drv].drops[locId] = 1;
    if (custNameForDrv) driverOrders[drv].customers[custNameForDrv] = (driverOrders[drv].customers[custNameForDrv] || 0) + 1;
    if (typeForDrv) driverOrders[drv].types[typeForDrv] = (driverOrders[drv].types[typeForDrv] || 0) + 1;
  });
  var driverList = Object.keys(driverOrders).map(function(name) {
    var custMap = driverOrders[name].customers || {};
    var typeMap = driverOrders[name].types || {};
    var topCustomers = Object.keys(custMap).sort(function(a,b){ return custMap[b]-custMap[a]; }).slice(0,3);
    var topTypes = Object.keys(typeMap).sort(function(a,b){ return typeMap[b]-typeMap[a]; });
    return { name:name, orders:driverOrders[name].orders, drops:Object.keys(driverOrders[name].drops).length, value:Math.round(driverOrders[name].value), isHired: /^hired driver/i.test(name), customers:topCustomers, types:topTypes };
  });
  var topDrivers = driverList.slice().sort(function(a,b) { return b.orders-a.orders; }).slice(0,5);
  // Order-count ranking hides drivers who carry only a few, very high-value deliveries
  // (e.g. a single route to a major supermarket) — surface those separately.
  var topDriversByValue = driverList.slice().sort(function(a,b) { return b.value-a.value; }).slice(0,5);

  // How much of today's dispatch relied on hired/agency drivers (no name in source data)
  // vs named in-house drivers — a bare phone number is the signal of a hired driver.
  var hiredDrivers = driverList.filter(function(d) { return d.isHired; });
  var inhouseDrivers = driverList.filter(function(d) { return !d.isHired; });
  var driverSourceSplit = {
    hired: {
      driver_count: hiredDrivers.length,
      orders: hiredDrivers.reduce(function(s,d){ return s+d.orders; }, 0),
      value: hiredDrivers.reduce(function(s,d){ return s+d.value; }, 0),
      drops: hiredDrivers.reduce(function(s,d){ return s+d.drops; }, 0)
    },
    inhouse: {
      driver_count: inhouseDrivers.length,
      orders: inhouseDrivers.reduce(function(s,d){ return s+d.orders; }, 0),
      value: inhouseDrivers.reduce(function(s,d){ return s+d.value; }, 0),
      drops: inhouseDrivers.reduce(function(s,d){ return s+d.drops; }, 0)
    },
    hired_driver_details: hiredDrivers.sort(function(a,b){ return b.value-a.value; })
  };

  // ── Own fleet vs 3PL drop split, and repeat-visit detection ──
  // 3PL orders are fulfilled by a third party (not our own fleet), so they're tracked
  // separately from our own-fleet drop count.
  var ownFleetDrops = 0, plDrops = 0;
  Object.keys(routes).forEach(function(r) {
    Object.keys(routes[r].locs).forEach(function(loc) {
      // A location counts as a "3PL drop" for this route only if EVERY visit to it
      // on this route was 3PL; otherwise it's counted as an own-fleet drop.
      if (locationVisits[loc] && locationVisits[loc].ownRoutes[r]) ownFleetDrops++;
      else plDrops++;
    });
  });

  var HIGH_VALUE_EXCEPTION_THRESHOLD = 100000;

  var repeatLocations = Object.keys(locationVisits)
    .map(function(loc) {
      var lv = locationVisits[loc];
      var ownRouteList = Object.keys(lv.ownRoutes);
      // For each own-fleet route that visited this location, what order type(s), value, order count,
      // and the actual order codes did it carry? (for full traceability back to source rows)
      var routeDetails = ownRouteList.map(function(r) {
        var typesHere = Object.keys(lv.typesByRoute[r] || {});
        return {
          route: r, types: typesHere,
          value: Math.round(lv.valueByRoute[r] || 0),
          order_count: lv.orderCountByRoute[r] || 0,
          orders: lv.ordersByRoute[r] || []
        };
      });
      var totalValue = routeDetails.reduce(function(s, rd) { return s + rd.value; }, 0);
      // All distinct types seen across all routes at this location
      var allTypes = {};
      routeDetails.forEach(function(rd) { rd.types.forEach(function(t) { allTypes[t] = true; }); });
      var distinctTypeCount = Object.keys(allTypes).length;
      // Legitimate split = different routes carried genuinely different order types
      // (e.g. one route Food, another Non-Food) — those must use separate trucks.
      // A real avoidable duplicate = multiple routes carrying the SAME type to the same address.
      var isLegitimateSplit = distinctTypeCount > 1;
      // High-value exception: an "avoidable" duplicate over AED 100k is more likely a genuinely
      // large order that needed splitting across trucks for capacity — flag for manual review
      // rather than assuming it's a routing mistake.
      var isHighValueException = !isLegitimateSplit && totalValue > HIGH_VALUE_EXCEPTION_THRESHOLD;
      return {
        location_id: loc,
        address: lv.address,
        customer: lv.customer,
        isCashOrder: lv.isCashOrder,
        isInternalHub: lv.isInternalHub,
        routes: ownRouteList,
        route_types: routeDetails,
        total_value: totalValue,
        visit_count: ownRouteList.length,
        is_legitimate_split: isLegitimateSplit,
        is_high_value_exception: isHighValueException,
        reason: isLegitimateSplit
          ? 'Different order types (' + Object.keys(allTypes).join(' + ') + ') — separate trucks required'
          : (isHighValueException
              ? 'EXCEPTION: AED ' + totalValue.toLocaleString() + ' — likely a genuinely large order needing capacity split, verify before flagging as routing error'
              : 'Same order type visited ' + ownRouteList.length + 'x — likely avoidable')
      };
    })
    .filter(function(l) { return l.visit_count > 1 && !l.isCashOrder && !l.isInternalHub; })
    .sort(function(a, b) { return (b.total_value - a.total_value) || (b.visit_count - a.visit_count); });

  var repeatLocationAvoidableCount = repeatLocations.filter(function(l) { return !l.is_legitimate_split; }).length;

  var cityTypeCrossOut = {};
  Object.keys(cityTypeCross).forEach(function(c) {
    cityTypeCrossOut[c] = cityTypeCross[c];
  });

  return {
    total_orders: totalOrders, total_value: Math.round(totalValue),
    total_routes: Object.keys(routes).length,
    total_drivers: Object.keys(driverOrders).length || Object.keys(driverSet).length,
    total_drops: Object.keys(routes).reduce(function(s,r){ return s + Object.keys(routes[r].locs).length; }, 0),
    food_orders: foodOrders, food_value: Math.round(foodValue),
    non_food_orders: nonFoodOrders, non_food_value: Math.round(nonFoodValue),
    pl_orders: plOrders, van_orders: vanOrders,
    frozen_orders: frozenOrders, frozen_value: Math.round(frozenValue),
    ambient_orders: ambientOrders, ambient_value: Math.round(ambientValue),
    type_breakdown: {
      DCV: { orders:orgStats.DCV.o, value:Math.round(orgStats.DCV.v) },
      DCF: { orders:orgStats.DCF.o, value:Math.round(orgStats.DCF.v) },
      DGC: { orders:orgStats.DGC.o, value:Math.round(orgStats.DGC.v) },
      DGS: { orders:orgStats.DGS.o, value:Math.round(orgStats.DGS.v) },
      DSN: { orders:orgStats.DSN.o, value:Math.round(orgStats.DSN.v) },
      DPS: { orders:orgStats.DPS.o, value:Math.round(orgStats.DPS.v) },
      DPB: { orders:orgStats.DPB.o, value:Math.round(orgStats.DPB.v) },
      HCP: { orders:orgStats.HCP.o, value:Math.round(orgStats.HCP.v) }
    },
    by_city: byCity, top_customers: topCustomers,
    top_drivers: topDrivers, top_drivers_by_value: topDriversByValue, driver_source_split: driverSourceSplit, top_routes: topRoutes,
    city_type_cross: cityTypeCrossOut,
    drops_by_city: dropsByCity,
    truck_cost_estimate: (function(){
      var unmatchedTruckTypes = {};
      var byType = {};       // label -> { rate, drop_count, vehicles:{} }
      var byRegion = {};     // city -> { label -> { rate, drop_count } }
      var byFoodType = {};   // 'Food' | 'Non-Food' | 'Mixed (Partition)' | 'Other' -> { drop_count, estimated_cost }
      var byTempFood = {};   // 'Frozen · Food' | 'Ambient · Non-Food' etc -> { drop_count, estimated_cost }

      // Pre-pass: how many drops does each vehicle make today? Needed to decide Multi vs Bulk
      // for the Vehicle Master fallback (transport team's rule: 1 drop = Bulk, 2+ = Multi).
      var dropCountByVehicle = {};
      Object.keys(dropRecords).forEach(function(key){
        var vid = dropRecords[key].vehicleId;
        if (vid) dropCountByVehicle[vid] = (dropCountByVehicle[vid] || 0) + 1;
      });

      Object.keys(dropRecords).forEach(function(key){
        var d = dropRecords[key];
        var rateEntry = null;
        if (d.truckType) {
          rateEntry = matchTruckRate(d.truckType);
        } else if (d.vehicleId) {
          // No truck-type text on this row — fall back to the Vehicle Master lookup by plate number.
          rateEntry = matchRateViaVehicleMaster(d.vehicleId, dropCountByVehicle[d.vehicleId] || 1);
        } else {
          return; // neither truck-type nor vehicle plate available — genuinely no info, not guessed
        }
        if (!rateEntry) { var uk = d.truckType || ('Vehicle ' + d.vehicleId); unmatchedTruckTypes[uk] = (unmatchedTruckTypes[uk] || 0) + 1; return; }

        if (!byType[rateEntry.label]) byType[rateEntry.label] = { rate: rateEntry.rate, drop_count: 0, vehicles: {} };
        byType[rateEntry.label].drop_count++;
        if (d.vehicleId) byType[rateEntry.label].vehicles[d.vehicleId] = 1;

        var city = d.city || 'Unknown';
        if (!byRegion[city]) byRegion[city] = {};
        if (!byRegion[city][rateEntry.label]) byRegion[city][rateEntry.label] = { rate: rateEntry.rate, drop_count: 0 };
        byRegion[city][rateEntry.label].drop_count++;

        var hasFood = !!d.types['food'];
        var hasNonFood = !!d.types['nonfood'];
        var foodCategory = (hasFood && hasNonFood) ? 'Mixed (Partition)' : (hasFood ? 'Food' : (hasNonFood ? 'Non-Food' : 'Other (3PL/Van)'));
        if (!byFoodType[foodCategory]) byFoodType[foodCategory] = { drop_count: 0, estimated_cost: 0 };
        byFoodType[foodCategory].drop_count++;
        byFoodType[foodCategory].estimated_cost += rateEntry.rate;

        var tempCategory = rateEntry.label.indexOf('Frozen') === 0 ? 'Frozen' : (rateEntry.label.indexOf('Ambient') === 0 ? 'Ambient' : 'Other');
        var tempFoodKey = tempCategory + ' · ' + foodCategory;
        if (!byTempFood[tempFoodKey]) byTempFood[tempFoodKey] = { drop_count: 0, estimated_cost: 0 };
        byTempFood[tempFoodKey].drop_count++;
        byTempFood[tempFoodKey].estimated_cost += rateEntry.rate;
      });

      var byTypeArr = Object.keys(byType).map(function(label){
        var d = byType[label];
        return { label: label, rate: d.rate, drop_count: d.drop_count, vehicle_count: Object.keys(d.vehicles).length, estimated_cost: d.drop_count * d.rate };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var byRegionArr = Object.keys(byRegion).map(function(city){
        var types = Object.keys(byRegion[city]).map(function(label){
          var d = byRegion[city][label];
          return { label: label, rate: d.rate, drop_count: d.drop_count, estimated_cost: d.drop_count * d.rate };
        }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });
        var regionCost = types.reduce(function(s,t){ return s + t.estimated_cost; }, 0);
        var regionDrops = types.reduce(function(s,t){ return s + t.drop_count; }, 0);
        return { city: city, types: types, total_cost: regionCost, total_drops: regionDrops };
      }).sort(function(a,b){ return b.total_cost - a.total_cost; });

      var byFoodTypeArr = Object.keys(byFoodType).map(function(cat){
        return { category: cat, drop_count: byFoodType[cat].drop_count, estimated_cost: byFoodType[cat].estimated_cost };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var byTempFoodArr = Object.keys(byTempFood).map(function(cat){
        return { category: cat, drop_count: byTempFood[cat].drop_count, estimated_cost: byTempFood[cat].estimated_cost };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var totalCost = byTypeArr.reduce(function(s,t){ return s + t.estimated_cost; }, 0);
      var totalVehicles = byTypeArr.reduce(function(s,t){ return s + t.vehicle_count; }, 0);
      var totalDropsBilled = byTypeArr.reduce(function(s,t){ return s + t.drop_count; }, 0);

      return {
        available: byTypeArr.length > 0,
        by_type: byTypeArr,
        by_region: byRegionArr,
        by_food_type: byFoodTypeArr,
        by_temp_food_type: byTempFoodArr,
        total_estimated_cost: totalCost,
        total_vehicles: totalVehicles,
        total_drops_billed: totalDropsBilled,
        unmatched_truck_types: unmatchedTruckTypes
      };
    })(),
    cost_analysis: {
      own_fleet_drops: ownFleetDrops,
      pl_drops: plDrops,
      repeat_location_count: repeatLocations.length,
      repeat_location_avoidable_count: repeatLocationAvoidableCount
    },
    repeat_locations: repeatLocations
  };
}

//  ORDER-LEVEL EXTRACTION FOR RE-DELIVERY TRACKING 
// Extracts a lightweight per-row record (order code, customer, value, route, org, type)
// from a dispatch file, used to detect the SAME order code appearing again on a LATER
// dispatch date (a true re-delivery — failed first attempt, re-attempted later), as
// opposed to a same-day duplicate.
function extractOrderRows(buffer) {
  var wb = XLSX.read(buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false, cellFormula: false });
  var sheetName = findDataSheet(wb);
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  if (!rows.length) return [];

  function findCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
    }) || null;
  }
  // Exact header match, same fallback-only role as in parseDispatch: "ORDER CODE" is
  // confirmed stable across days for the same re-delivered order (verified against real
  // files), so it's preferred. Plain "ORDER" is only used if "ORDER CODE" is missing.
  function findExactCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.trim().toUpperCase() === n.toUpperCase(); });
    }) || null;
  }
  var C = {
    orderCode: findCol('ORDER CODE', 'ORDER_CODE') || findExactCol('ORDER', 'ORDER NUMBER', 'ORDER_NUMBER') || findCol('ORDER '),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount: findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    route: findCol('ROUTE'),
    org: findCol('ORG') || findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU'),
    type: findCol('TYPE'),
    temperature: findCol('TEMPERATURE')
  };
  console.log('Re-delivery tracking cols:', JSON.stringify(C));
  if (!C.orderCode) return []; // no order code column in this file — can't track re-delivery

  var out = [];
  rows.forEach(function(row) {
    var code = toStr(row[C.orderCode]);
    if (!code) return;
    out.push({
      order_code: code,
      customer: C.customer ? toStr(row[C.customer]) : '',
      value: C.amount ? (parseFloat(row[C.amount]) || 0) : 0,
      route: C.route ? toStr(row[C.route]) : '',
      org: C.org ? toStr(row[C.org]).toUpperCase() : '',
      drop_type: C.type ? normaliseType(row[C.type]) : '',
      temperature: C.temperature ? toStr(row[C.temperature]).trim() : ''
    });
  });
  return out;
}

async function saveOrderTracking(dateKey, orderRows) {
  try {
    await pool.query('DELETE FROM order_tracking WHERE date_key=$1', [dateKey]);
    if (!orderRows.length) return true;
    var CHUNK = 500;
    for (var i = 0; i < orderRows.length; i += CHUNK) {
      var chunk = orderRows.slice(i, i + CHUNK);
      var vals = [];
      var phs = [];
      var idx = 1;
      chunk.forEach(function(r) {
        phs.push('($' + idx + ',$' + (idx+1) + ',$' + (idx+2) + ',$' + (idx+3) + ',$' + (idx+4) + ',$' + (idx+5) + ',$' + (idx+6) + ',$' + (idx+7) + ')');
        vals.push(r.order_code, dateKey, r.customer, r.value, r.route, r.org, r.drop_type, r.temperature);
        idx += 8;
      });
      await pool.query('INSERT INTO order_tracking (order_code, date_key, customer, value, route, org, drop_type, temperature) VALUES ' + phs.join(','), vals);
    }
    return true;
  } catch(e) {
    console.error('saveOrderTracking error:', e.message);
    return false;
  }
}

//  DISPATCH MEMORY (+ DB) 
var dispatchHistory = {};
var currentDispatch = null;

async function loadDispatchFromDB() {
  try {
    var rows = await dbLoadDispatch();
    if (rows.length > 0) {
      rows.forEach(function(r) {
        dispatchHistory[r.date_key] = {
          uploadedAt: r.uploaded_at, uploadedBy: r.uploaded_by,
          summary: r.summary, csvText: r.csv_text, date: r.date_key
        };
      });
      var latest = rows[0];
      currentDispatch = dispatchHistory[latest.date_key];
      console.log('Loaded', rows.length, 'dispatch dates from DB');
      return true;
    }
  } catch(e) { console.error('loadDispatchFromDB error:', e.message); }
  // Fallback to file
  try {
    var saved = loadJSON(DISPATCH_FILE);
    if (saved) {
      dispatchHistory = saved.history || {};
      var keys = Object.keys(dispatchHistory).sort().reverse();
      if (keys.length) currentDispatch = dispatchHistory[keys[0]];
      console.log('Loaded dispatch from file:', keys.length, 'dates');
    }
  } catch(e) { console.error('loadDispatchFromDB file fallback error:', e.message); }
  return false;
}
loadDispatchFromDB();

app.post('/api/dispatch/upload', upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    var summary = parseDispatch(req.file.buffer);
    if (!summary) return res.status(400).json({ error: 'Could not parse file' });
    var dateKey    = req.body.dateKey    || new Date().toISOString().split('T')[0];
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var wb2 = XLSX.read(req.file.buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false });
    var csv = XLSX.utils.sheet_to_csv(wb2.Sheets[findDataSheet(wb2)]);
    // Save to DB
    var dbOk = await dbSaveDispatch(dateKey, uploadedBy, summary, csv);
    // Also save to file as backup
    var entry = { uploadedAt:new Date().toISOString(), uploadedBy:uploadedBy, csvText:csv.substring(0,200000), summary:summary, date:dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;
    // Keep up to 180 days (6 months)
    var keys = Object.keys(dispatchHistory).sort();
    while (keys.length > 180) delete dispatchHistory[keys.shift()];
    saveJSON(DISPATCH_FILE, { history:dispatchHistory });
    console.log('Dispatch saved:', dateKey, dbOk ? '(DB+file)' : '(file only)');
    // Track order codes for this date so re-delivery (same order, later day) can be detected.
    // Wrapped separately so any issue here never breaks the main dispatch upload response.
    try {
      var orderRows = extractOrderRows(req.file.buffer);
      await saveOrderTracking(dateKey, orderRows);
      console.log('Order tracking saved:', dateKey, orderRows.length, 'order rows');
    } catch(trackErr) {
      console.error('Order tracking error (non-fatal):', trackErr.message);
    }
    res.json({ success:true, summary:summary, uploadedAt:entry.uploadedAt, date:dateKey });
  } catch(e) {
    console.error('Dispatch upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.get('/api/dispatch/status', function(req, res) {
  var avail = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData:false, availableDates:avail });
  res.json({ hasData:true, uploadedAt:currentDispatch.uploadedAt, uploadedBy:currentDispatch.uploadedBy, summary:currentDispatch.summary, date:currentDispatch.date, availableDates:avail });
});

app.get('/api/dispatch/date/:dateKey', function(req, res) {
  var entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData:false });
  currentDispatch = entry;
  res.json({ hasData:true, uploadedAt:entry.uploadedAt, uploadedBy:entry.uploadedBy, summary:entry.summary, date:entry.date });
});

// Re-delivery tracking: finds order codes active on the given date that ALSO appear on
// any EARLIER dispatch date — i.e. the same order was dispatched before, presumably
// failed, and is being re-delivered now. Same-day duplicates are not counted here.
app.get('/api/dispatch/redelivery/:dateKey', async function(req, res) {
  try {
    var dateKey = req.params.dateKey;
    var todayRes = await pool.query('SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1', [dateKey]);
    var todayCodes = todayRes.rows.map(function(r){ return r.order_code; });
    if (!todayCodes.length) return res.json({ hasData:true, dateKey:dateKey, total_repeated_orders:0, total_value_at_risk:0, orders:[] });

    var histRes = await pool.query(
      'SELECT order_code, date_key, customer, value, route, org, drop_type, temperature FROM order_tracking WHERE order_code = ANY($1) AND date_key <= $2 ORDER BY date_key ASC',
      [todayCodes, dateKey]
    );

    var byCode = {};
    histRes.rows.forEach(function(r) {
      if (!byCode[r.order_code]) byCode[r.order_code] = [];
      byCode[r.order_code].push(r);
    });

    var repeated = [];
    Object.keys(byCode).forEach(function(code) {
      var occ = byCode[code];
      var distinctDates = Array.from(new Set(occ.map(function(o) {
        var dk = o.date_key;
        return (dk && dk.toISOString) ? dk.toISOString().split('T')[0] : String(dk).split('T')[0];
      })));
      if (distinctDates.length > 1) {
        var latest = occ[occ.length - 1];
        repeated.push({
          order_code: code,
          customer: latest.customer || '',
          value: parseFloat(latest.value) || 0,
          org: latest.org || '',
          route: latest.route || '',
          drop_type: latest.drop_type || '',
          temperature: latest.temperature || '',
          times_delivered: distinctDates.length,
          dates: distinctDates
        });
      }
    });
    repeated.sort(function(a, b) { return b.value - a.value; });

    var totalValue = repeated.reduce(function(s, r) { return s + r.value; }, 0);
    res.json({
      hasData: true,
      dateKey: dateKey,
      total_repeated_orders: repeated.length,
      total_value_at_risk: Math.round(totalValue),
      orders: repeated.slice(0, 200)
    });
  } catch(e) {
    console.error('redelivery endpoint error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ============================================================
// TRANSPORT COST RECONCILIATION — cross-checks the transport team's own
// reported drop counts (their raw "Distinct DROP ID" export) against what
// this app captured from the daily dispatch uploads, so discrepancies show
// up before the monthly invoice arrives rather than after paying it.
// ============================================================

function normalizeReconcileHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
function matchReconcileField(header) {
  var h = normalizeReconcileHeader(header);
  if (h.indexOf('dropid') !== -1) return 'drop_id';
  if (h.indexOf('taskstatus') !== -1) return 'task_status';
  if (h.indexOf('taskid') !== -1) return 'task_order_id';
  if (h.indexOf('finaldropsremark') !== -1) return 'drop_class';
  if (h.indexOf('operatingunit') !== -1) return 'operating_unit';
  if (h.indexOf('dispatchdate') !== -1 || h === 'date') return 'dispatch_date';
  return null;
}
// Transport's TASK ID sometimes carries a re-attempt suffix like
// "-2026-07-01-1" that the AKI dispatch file's ORDER CODE never has —
// stripping it is what lets the two systems' order identifiers line up
// (confirmed: 923 of 930 real orders matched exactly once stripped).
function stripTaskIdSuffix(v) {
  return String(v || '').trim().replace(/-\d{4}-\d{2}-\d{2}(-\d+)?$/, '');
}

app.post('/api/dispatch/reconcile/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    var wb;
    try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
    catch (e) { return res.status(400).json({ error: 'Could not read that file.' }); }

    // Find the sheet that actually has DROP ID + TASK STATUS + Final Drops Remarks +
    // a date column — the transport team's export has multiple sheets (Summary,
    // Bulk, Multi pivots) and only "Raw Data" has what's needed here.
    var bestSheetName = wb.SheetNames[0], bestScore = -1;
    wb.SheetNames.forEach(function (name) {
      var s = wb.Sheets[name];
      var firstRow = XLSX.utils.sheet_to_json(s, { header: 1, defval: '' })[0] || [];
      var score = 0, hasDropId = false, hasStatus = false, hasClass = false, hasDate = false;
      firstRow.forEach(function (cell) {
        var f = matchReconcileField(cell);
        if (f) { score++; if (f === 'drop_id') hasDropId = true; if (f === 'task_status') hasStatus = true; if (f === 'drop_class') hasClass = true; if (f === 'dispatch_date') hasDate = true; }
      });
      if (hasDropId && hasStatus && hasClass && hasDate && score > bestScore) { bestScore = score; bestSheetName = name; }
    });
    if (bestScore === -1) return res.status(400).json({ error: 'Could not find a sheet with DROP ID, TASK STATUS, Final Drops Remarks, and a date column. This should be the "Raw Data" sheet from the transport team\'s export.' });

    var sheet = wb.Sheets[bestSheetName];
    var rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (!rawRows.length) return res.status(400).json({ error: 'The file has no data rows.' });

    var fieldByHeader = {};
    Object.keys(rawRows[0]).forEach(function (h) { var f = matchReconcileField(h); if (f) fieldByHeader[h] = f; });

    // TASK_STATUS_RULE documents the assumption in effect — currently COMPLETED-only,
    // matching how a transport-team invoice should only bill for completed drops.
    // If Azhar confirms a different rule with the transport team, this is the one
    // line to change.
    // Per Azhar's direction: a dispatched drop is a dispatched drop, whatever
    // the transport team's internal TASK STATUS says — that status describes
    // THEIR workflow, not whether the trip happened. So the primary count
    // below includes every status. COMPLETED is tracked in parallel purely as
    // supplementary "how much of this day is closed out" context.
    var COMPLETED_STATUS = 'COMPLETED';

    var byDateClass = {}; // "date|class" -> Set of drop_ids (ALL statuses — this is the real count)
    var completedByDateClass = {}; // "date|class" -> Set of drop_ids where status === COMPLETED (context only)
    var orderIdsByDate = {}; // date -> Set of stripped TASK IDs, for order-level diffing
    var totalRows = 0, skippedNonConsumer = 0;
    rawRows.forEach(function (raw) {
      var rec = {};
      Object.keys(raw).forEach(function (h) { var f = fieldByHeader[h]; if (f) rec[f] = raw[h]; });
      if (!rec.drop_id || !rec.dispatch_date) return;
      // Only filter by OPERATING UNIT when that column actually exists in this file — a
      // Consumer-only export (no such column) is trusted as-is. A full multi-department
      // export gets filtered down to Consumer here, so uploading either kind of file
      // produces the same Consumer-only reconciliation.
      if (rec.operating_unit !== undefined && String(rec.operating_unit).trim().toUpperCase() !== 'CONSUMER') {
        skippedNonConsumer++;
        return;
      }
      totalRows++;
      var dateStr = toDateStrGeneric(rec.dispatch_date);
      if (!dateStr) return;
      var cls = String(rec.drop_class || 'Other').trim() || 'Other';
      var key = dateStr + '|' + cls;
      if (!byDateClass[key]) byDateClass[key] = new Set();
      byDateClass[key].add(String(rec.drop_id));
      var status = String(rec.task_status || '').trim().toUpperCase();
      if (status === COMPLETED_STATUS) {
        if (!completedByDateClass[key]) completedByDateClass[key] = new Set();
        completedByDateClass[key].add(String(rec.drop_id));
      }
      if (rec.task_order_id) {
        var stripped = stripTaskIdSuffix(rec.task_order_id);
        if (stripped) {
          if (!orderIdsByDate[dateStr]) orderIdsByDate[dateStr] = new Set();
          orderIdsByDate[dateStr].add(stripped);
        }
      }
    });

    var batchId = 'RECON-' + Date.now();
    var upserted = 0;
    for (var key in byDateClass) {
      var parts = key.split('|');
      var dateStr = parts[0], cls = parts[1];
      var count = byDateClass[key].size;
      var completedCount = completedByDateClass[key] ? completedByDateClass[key].size : 0;
      await pool.query(
        `INSERT INTO transport_drop_reconciliation (date_key, drop_class, reported_drops, completed_drops, task_status_rule, upload_batch_id, uploaded_by)
         VALUES ($1,$2,$3,$4,'ALL',$5,$6)
         ON CONFLICT (date_key, drop_class) DO UPDATE SET reported_drops=$3, completed_drops=$4, task_status_rule='ALL', upload_batch_id=$5, uploaded_by=$6, uploaded_at=NOW()`,
        [dateStr, cls, count, completedCount, batchId, req.user ? req.user.username : 'Admin']
      );
      upserted++;
    }

    var orderIdsSaved = 0;
    for (var od in orderIdsByDate) {
      // Clear this date's previous order IDs first so a re-upload doesn't leave stale entries.
      await pool.query('DELETE FROM transport_order_ids WHERE date_key=$1', [od]);
      var ids = Array.from(orderIdsByDate[od]);
      for (var i = 0; i < ids.length; i += 500) {
        var chunk = ids.slice(i, i + 500);
        var values = chunk.map(function (_, ci) { return '($1,$' + (ci + 2) + ',$' + (chunk.length + 2) + ')'; }).join(',');
        await pool.query(
          'INSERT INTO transport_order_ids (date_key, order_id, upload_batch_id) VALUES ' + values + ' ON CONFLICT (date_key, order_id) DO NOTHING',
          [od].concat(chunk).concat([batchId])
        );
      }
      orderIdsSaved += ids.length;
    }

    res.json({
      batch_id: batchId,
      total_rows: totalRows,
      task_status_rule: 'ALL (every dispatched drop, any status)',
      date_class_buckets_saved: upserted,
      order_ids_saved: orderIdsSaved,
      order_level_diff_available: orderIdsSaved > 0,
      skipped_non_consumer: skippedNonConsumer,
      sheet_used: bestSheetName
    });
  } catch (e) {
    console.error('reconcile upload error:', e.message);
    res.status(500).json({ error: 'Upload failed: ' + e.message });
  }
});

app.get('/api/dispatch/reconcile', requireAuth, async function (req, res) {
  try {
    var params = [], clauses = [];
    if (req.query.date_from) { params.push(req.query.date_from); clauses.push('date_key >= $' + params.length); }
    if (req.query.date_to) { params.push(req.query.date_to); clauses.push('date_key <= $' + params.length); }
    var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';

    var transportRows = await pool.query(
      `SELECT date_key::text, drop_class, reported_drops, completed_drops, task_status_rule, uploaded_at FROM transport_drop_reconciliation ${where} ORDER BY date_key`, params
    );
    if (!transportRows.rows.length) return res.json({ has_data: false, days: [] });

    var byDate = {};
    transportRows.rows.forEach(function (r) {
      if (!byDate[r.date_key]) byDate[r.date_key] = { date: r.date_key, transport_total: 0, completed_total: 0, by_class: {}, uploaded_at: r.uploaded_at };
      byDate[r.date_key].transport_total += r.reported_drops;
      byDate[r.date_key].completed_total += r.completed_drops;
      byDate[r.date_key].by_class[r.drop_class] = r.reported_drops;
    });

    var dateKeys = Object.keys(byDate).sort();
    var appRows = await pool.query(
      `SELECT date_key::text, summary FROM dispatch_data WHERE date_key = ANY($1::date[])`, [dateKeys]
    );
    var appByDate = {};
    appRows.rows.forEach(function (r) { appByDate[r.date_key] = r.summary; });

    var days = dateKeys.map(function (d) {
      var t = byDate[d];
      var appSummary = appByDate[d] || null;
      var appDrops = appSummary && appSummary.total_drops !== undefined ? appSummary.total_drops : null;
      var variance = (appDrops !== null) ? (appDrops - t.transport_total) : null;
      var variancePct = (appDrops !== null && t.transport_total > 0) ? +((variance / t.transport_total) * 100).toFixed(1) : null;
      var completedPct = t.transport_total > 0 ? +((t.completed_total / t.transport_total) * 100).toFixed(1) : null;
      return {
        date: d,
        app_captured: appDrops,
        app_has_data: appSummary !== null,
        transport_reported: t.transport_total,
        transport_by_class: t.by_class,
        transport_completed: t.completed_total,
        transport_completed_pct: completedPct,
        variance: variance,
        variance_pct: variancePct
      };
    });

    var totalApp = days.reduce(function (s, d) { return s + (d.app_captured || 0); }, 0);
    var totalTransport = days.reduce(function (s, d) { return s + d.transport_reported; }, 0);
    var daysMissingAppData = days.filter(function (d) { return !d.app_has_data; }).length;

    res.json({
      has_data: true,
      days: days,
      total_app_captured: totalApp,
      total_transport_reported: totalTransport,
      total_variance: totalApp - totalTransport,
      days_missing_app_data: daysMissingAppData
    });
  } catch (e) {
    console.error('reconcile compare error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Order-level diff: which specific orders does the app have that transport
// doesn't (and vice versa) for one date — the actionable detail behind a
// day-level count mismatch. Matches on order_tracking (this app's own
// per-day order log, already populated from daily dispatch uploads) against
// transport_order_ids (populated from the transport file's TASK ID column).
app.get('/api/dispatch/reconcile/order-diff', requireAuth, async function (req, res) {
  try {
    var date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date is required, e.g. ?date=2026-07-18' });

    var appRows = await pool.query('SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1', [date]);
    var transportRows = await pool.query('SELECT order_id FROM transport_order_ids WHERE date_key=$1', [date]);

    if (!appRows.rows.length && !transportRows.rows.length) {
      return res.json({ has_data: false, date: date });
    }

    var appSet = new Set(appRows.rows.map(function (r) { return r.order_code; }));
    var transportSet = new Set(transportRows.rows.map(function (r) { return r.order_id; }));

    var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); });
    var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); });
    var matched = Array.from(appSet).filter(function (o) { return transportSet.has(o); });

    res.json({
      has_data: true,
      date: date,
      app_total: appSet.size,
      transport_total: transportSet.size,
      matched_count: matched.length,
      app_only: appOnly.sort(),
      transport_only: transportOnly.sort()
    });
  } catch (e) {
    console.error('order-diff error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/dispatch/reconcile/order-diff/export', requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date is required' });
    var appRows = await pool.query('SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1', [date]);
    var transportRows = await pool.query('SELECT order_id FROM transport_order_ids WHERE date_key=$1', [date]);
    var appSet = new Set(appRows.rows.map(function (r) { return r.order_code; }));
    var transportSet = new Set(transportRows.rows.map(function (r) { return r.order_id; }));
    var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); }).sort();
    var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); }).sort();

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();
    var ws = wb.addWorksheet('Order Diff ' + date);
    ws.columns = [{ width: 30 }, { width: 30 }];
    var titleRow = ws.addRow(['ORDER-LEVEL RECONCILIATION — ' + date]);
    ws.mergeCells('A' + titleRow.number + ':B' + titleRow.number);
    titleRow.font = { bold: true, size: 14 };
    ws.addRow(['App orders: ' + appSet.size, 'Transport orders: ' + transportSet.size]);
    ws.addRow([]);
    var hdr = ws.addRow(['Dispatched by AKI, missing from Transport file', 'In Transport file, missing from AKI dispatch']);
    hdr.font = { bold: true };
    var maxLen = Math.max(appOnly.length, transportOnly.length);
    for (var i = 0; i < maxLen; i++) {
      ws.addRow([appOnly[i] || '', transportOnly[i] || '']);
    }
    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Order_Diff_' + date + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});

app.get('/api/dispatch/reconcile/export', requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var params = [], clauses = [];
    if (req.query.date_from) { params.push(req.query.date_from); clauses.push('date_key >= $' + params.length); }
    if (req.query.date_to) { params.push(req.query.date_to); clauses.push('date_key <= $' + params.length); }
    var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
    var transportRows = await pool.query(`SELECT date_key::text, drop_class, reported_drops, completed_drops FROM transport_drop_reconciliation ${where} ORDER BY date_key`, params);
    var byDate = {};
    transportRows.rows.forEach(function (r) {
      if (!byDate[r.date_key]) byDate[r.date_key] = { total: 0, completed: 0, classes: {} };
      byDate[r.date_key].total += r.reported_drops;
      byDate[r.date_key].completed += r.completed_drops;
      byDate[r.date_key].classes[r.drop_class] = r.reported_drops;
    });
    var dateKeys = Object.keys(byDate).sort();
    var appRows = await pool.query(`SELECT date_key::text, summary FROM dispatch_data WHERE date_key = ANY($1::date[])`, [dateKeys]);
    var appByDate = {};
    appRows.rows.forEach(function (r) { appByDate[r.date_key] = r.summary; });

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();
    var ws = wb.addWorksheet('Reconciliation');
    ws.columns = [{ width: 14 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }];
    var titleRow = ws.addRow(['TRANSPORT DROP RECONCILIATION — App Captured vs Transport Reported (every dispatched drop, any status)']);
    ws.mergeCells('A' + titleRow.number + ':H' + titleRow.number);
    titleRow.font = { bold: true, size: 14 };
    ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    ws.addRow([]);
    var hdr = ws.addRow(['Date', 'App Captured', 'Transport Bulk', 'Transport Multi', 'Transport Total', 'Variance', 'Variance %', 'Completed So Far']);
    hdr.font = { bold: true };
    var totA = 0, totT = 0, totTBothSides = 0, daysBothSides = 0;
    dateKeys.forEach(function (d) {
      var appSummary = appByDate[d];
      var appDrops = appSummary && appSummary.total_drops !== undefined ? appSummary.total_drops : null;
      var t = byDate[d];
      var variance = appDrops !== null ? appDrops - t.total : null;
      var variancePct = (appDrops !== null && t.total > 0) ? +((variance / t.total) * 100).toFixed(1) : null;
      var completedPct = t.total > 0 ? Math.round((t.completed / t.total) * 100) : 0;
      totT += t.total;
      if (appDrops !== null) { totA += appDrops; totTBothSides += t.total; daysBothSides++; }
      ws.addRow([d, appDrops === null ? 'No app data' : appDrops, t.classes.Bulk || 0, t.classes.Multi || 0, t.total, variance === null ? '' : variance, variancePct === null ? '' : variancePct, completedPct + '% (' + t.completed + ' of ' + t.total + ')']);
    });
    ws.addRow([]);
    // A straight sum across ALL dates is misleading whenever some dates are missing an
    // app-side upload (their Transport total would count with nothing to offset it) — so
    // the meaningful total only covers dates where BOTH sides actually have data.
    var totalRow = ws.addRow(['TOTAL (days with BOTH App + Transport data, ' + daysBothSides + ' of ' + dateKeys.length + ' days)', totA, '', '', totTBothSides, totA - totTBothSides, '']);
    totalRow.font = { bold: true };
    var allDatesRow = ws.addRow(['All ' + dateKeys.length + ' days — Transport total only (App data incomplete)', '', '', '', totT, '', '']);
    allDatesRow.font = { italic: true, color: { argb: 'FF8FA0B0' } };

    // ---- Sheet 2: Order-Level Issues — ready to send straight to the transport team ----
    var orderIdRows = await pool.query('SELECT DISTINCT date_key::text FROM transport_order_ids ORDER BY date_key');
    var wsIssues = wb.addWorksheet('Order-Level Issues');
    wsIssues.columns = [{ width: 14 }, { width: 26 }, { width: 40 }];
    var issueTitle = wsIssues.addRow(['ORDER-LEVEL ISSUES — orders that don\'t match between AKI dispatch and the transport team\'s file']);
    wsIssues.mergeCells('A' + issueTitle.number + ':C' + issueTitle.number);
    issueTitle.font = { bold: true, size: 14 };
    wsIssues.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    wsIssues.addRow([]);
    var issueHdr = wsIssues.addRow(['Date', 'Order Code', 'Issue']);
    issueHdr.font = { bold: true };
    var totalIssueRows = 0, datesSkippedNoAppData = [];
    for (var i = 0; i < orderIdRows.rows.length; i++) {
      var dk = orderIdRows.rows[i].date_key;
      var appOrderRows = await pool.query('SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1', [dk]);
      if (!appOrderRows.rows.length) { datesSkippedNoAppData.push(dk); continue; } // can't diff a date with no app dispatch upload at all
      var transportOrderRows = await pool.query('SELECT order_id FROM transport_order_ids WHERE date_key=$1', [dk]);
      var appSet = new Set(appOrderRows.rows.map(function (r) { return r.order_code; }));
      var transportSet = new Set(transportOrderRows.rows.map(function (r) { return r.order_id; }));
      var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); }).sort();
      var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); }).sort();
      appOnly.forEach(function (o) {
        var row = wsIssues.addRow([dk, o, 'Dispatched by AKI — missing from transport file']);
        row.getCell(3).font = { color: { argb: 'FFE05C5C' } };
        totalIssueRows++;
      });
      transportOnly.forEach(function (o) {
        var row = wsIssues.addRow([dk, o, 'In transport file — missing from AKI dispatch (ask why this was added)']);
        row.getCell(3).font = { color: { argb: 'FFB8860B' }, bold: true };
        totalIssueRows++;
      });
    }
    if (!totalIssueRows) {
      wsIssues.addRow(['No order-level issues found for any date with data on both sides.']);
    }
    if (datesSkippedNoAppData.length) {
      wsIssues.addRow([]);
      wsIssues.addRow(['Dates skipped (no App dispatch upload for that date, so no order-level comparison possible):']);
      wsIssues.addRow([datesSkippedNoAppData.join(', ')]);
    }

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Transport_Reconciliation_' + Date.now() + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});

app.post('/api/dispatch/ask', function(req, res) {
  try {
    if (!currentDispatch) return res.json({ result:'No dispatch data. Please upload first.' });
    var s = currentDispatch.summary;
    var context = 'Date: '+currentDispatch.date+'\nTotal Orders: '+s.total_orders+'\nTotal Value: AED '+s.total_value+'\nFood: '+s.food_orders+' orders AED '+s.food_value+'\nNon-Food: '+s.non_food_orders+' orders AED '+s.non_food_value+'\n3PL: '+s.pl_orders+'\n\nCSV:\n'+(currentDispatch.csvText||'').substring(0,8000);
    anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:1500,
      messages:[{ role:'user', content:'You are AZHAR-AI Dispatch Intelligence for UAE logistics.\n\n'+context+'\n\nQuestion: '+req.body.question+'\n\nAnswer with exact numbers. Use AED for currency.' }]
    }).then(function(msg) { res.json({ result: msg.content[0].text }); })
      .catch(function(e) { res.status(500).json({ error: e.message }); });
  } catch(e) { if (!res.headersSent) res.status(500).json({ error: e.message }); }
});

//  REJECTION STORE (+ DB) 
var rejectionData = null;

async function loadRejectionFromDB() {
  try {
    var row = await dbLoadRejection();
    if (row) {
      // Safely parse orgs/months — may be object (JSONB) or string (old saves)
      var orgs = row.orgs;
      var months = row.months;
      if (typeof orgs === 'string') try { orgs = JSON.parse(orgs); } catch(e) { orgs = {}; }
      if (typeof months === 'string') try { months = JSON.parse(months); } catch(e) { months = {}; }
      // Version check: if 'all' org has no detail array, data is old — needs re-upload
      var hasDetail = orgs && orgs.all && Array.isArray(orgs.all.detail) && orgs.all.detail.length > 0;
      rejectionData = {
        uploadedAt: row.uploaded_at, uploadedBy: row.uploaded_by,
        fileName: row.file_name, totalOrders: row.total_orders,
        orgs: orgs, months: months,
        needsReupload: !hasDetail
      };
      console.log('Loaded rejection from DB. hasDetail:', hasDetail, 'orgs keys:', Object.keys(orgs||{}).length);
      return true;
    }
  } catch(e) { console.error('loadRejectionFromDB error:', e.message); }
  try {
    var saved = loadJSON(REJECTION_FILE);
    if (saved) { rejectionData = saved; console.log('Loaded rejection from file'); }
  } catch(e) { console.error('loadRejectionFromDB file fallback error:', e.message); }
  return false;
}
loadRejectionFromDB();

app.post('/api/rejection/upload', upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error:'No file received' });
    var ext = path.extname(req.file.originalname||'').toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls' && ext !== '.csv')
      return res.status(400).json({ error:'Please upload .xlsx, .xls or .csv' });

    console.log('Reading rejection file:', req.file.originalname, req.file.size, 'bytes');
    var rows = [];
    if (ext === '.csv') {
      // Try UTF-8 first, fall back to latin1 for special characters
      var csvText;
      try {
        csvText = req.file.buffer.toString('utf8');
        // Check for replacement characters indicating wrong encoding
        if (csvText.includes('\uFFFD')) {
          csvText = req.file.buffer.toString('latin1');
          console.log('CSV: switched to latin1 encoding');
        }
      } catch(e) {
        csvText = req.file.buffer.toString('latin1');
        console.log('CSV: using latin1 encoding');
      }
      var csvRows = csvText.split('\n').filter(function(l){return l.trim();});
      if (csvRows.length < 2) return res.status(400).json({ error:'CSV file is empty' });
      function parseCSVLine(line) {
        var result=[], cell='', inQ=false;
        for (var ci=0; ci<line.length; ci++) {
          var ch=line[ci];
          if(ch==='"'){inQ=!inQ;}
          else if(ch===','&&!inQ){result.push(cell.trim());cell='';}
          else{cell+=ch;}
        }
        result.push(cell.trim());
        return result;
      }
      var headers = parseCSVLine(csvRows[0]).map(function(h){return h.replace(/"/g,'').trim();});
      for (var ci=1; ci<csvRows.length; ci++) {
        if (!csvRows[ci].trim()) continue;
        var vals = parseCSVLine(csvRows[ci]);
        var rowObj = {};
        headers.forEach(function(h,hi){ rowObj[h] = (vals[hi]||'').replace(/"/g,'').trim(); });
        rows.push(rowObj);
      }
      console.log('CSV rows parsed:', rows.length);
    } else {
      var wb = XLSX.read(req.file.buffer, { type:'buffer', dense:true, cellDates:false, cellNF:false, cellHTML:false, cellFormula:false });
      var sheetName = findDataSheet(wb);
      var ws = wb.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json(ws, { defval:'', raw:true });
      console.log('Excel rows:', rows.length, 'sheet:', sheetName);
    }
    if (!rows.length) return res.status(400).json({ error:'No rows found' });

    var keys0 = Object.keys(rows[0]);
    function findC() {
      var names = Array.prototype.slice.call(arguments);
      return keys0.find(function(k) {
        return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
      }) || null;
    }
    var RC = {
      status:  findC('FINAL STATUS', 'STATUS'),
      org:     findC('ORGANIZATION') || findC('ORG-BU'),
      date:    findC('D DATE', 'DATE', 'DELIVERY DATE'),
      root:    findC('FINAL- ROOT', 'FINA- ROOT', 'ROOT CAUSE', 'ROOT_CAUSE', 'REASON-1'),
      cust:    findC('CUSTOMER NAME', 'CUSTOMER'),
      addr:    findC('FULL ADDRESS', 'ADDRESS MATCHING', 'ADDRESS'),
      area:    findC('AREA', 'CITY'),
      value:   findC('VALUE', 'AMOUNT'),
      type:    findC('TYPE'),
      source:  findC('REMAKE -3', 'REMAKE') || findC('INTERNAL/EXTERNAL'),
      orderNo: findC('ORDER NO', 'ORDER_NO', 'ORDERNO', 'SHIPMENT_ID')
    };
    console.log('Rejection cols:', JSON.stringify(RC));

    function isRej(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1==='REJECTION'||s1==='REJECTED'||s2==='R/D'||s2==='HOLD'||s2==='RD';
    }
    function isDel(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1==='DELIVERED'||s1.includes('DELIVER')||s2.includes('DELIVER')||s2==='D';
    }
    function parseDate(v) {
      if (!v) return null;
      if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

      // Excel serial date number (e.g. 46000)
      if (typeof v === 'number') {
        try { var unix = Math.round((v - 25569) * 86400 * 1000); var dd = new Date(unix); if (!isNaN(dd.getTime())) return dd; } catch(e2) {}
        return null;
      }

      var s = String(v).trim();
      if (!s) return null;

      var MONTH_NAMES = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

      // Format: "13-Apr-26" or "13-Apr-2026" or "13 Apr 2026" (day - month name - year)
      var m1 = s.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3,})[\s\-\/](\d{2,4})$/);
      if (m1) {
        var mon1 = MONTH_NAMES[m1[2].toLowerCase().substring(0,3)];
        if (mon1 !== undefined) {
          var yr1 = parseInt(m1[3], 10); if (yr1 < 100) yr1 += 2000;
          var dt1 = new Date(yr1, mon1, parseInt(m1[1], 10));
          if (!isNaN(dt1.getTime())) return dt1;
        }
      }

      // Format: "Apr-13-26" or "Apr 13 2026" (month name - day - year)
      var m2 = s.match(/^([A-Za-z]{3,})[\s\-\/](\d{1,2})[\s\-\/](\d{2,4})$/);
      if (m2) {
        var mon2 = MONTH_NAMES[m2[1].toLowerCase().substring(0,3)];
        if (mon2 !== undefined) {
          var yr2 = parseInt(m2[3], 10); if (yr2 < 100) yr2 += 2000;
          var dt2 = new Date(yr2, mon2, parseInt(m2[2], 10));
          if (!isNaN(dt2.getTime())) return dt2;
        }
      }

      // Format: "YYYY-MM-DD" (ISO)
      var m3 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m3) {
        var dt3 = new Date(parseInt(m3[1],10), parseInt(m3[2],10)-1, parseInt(m3[3],10));
        if (!isNaN(dt3.getTime())) return dt3;
      }

      // Format: "M/D/YYYY" or "D/M/YYYY" (slash-separated, ambiguous — resolved below)
      var m4 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (m4) {
        var a = parseInt(m4[1],10), b = parseInt(m4[2],10), yr4 = parseInt(m4[3],10);
        if (yr4 < 100) yr4 += 2000;
        var month4, day4;
        if (a > 12 && b <= 12) { day4 = a; month4 = b; }       // first number can't be a month -> D/M/Y
        else { month4 = a; day4 = b; }                         // default: M/D/Y (matches the real data seen)
        var dt4 = new Date(yr4, month4-1, day4);
        if (!isNaN(dt4.getTime())) return dt4;
      }

      // Format: "13.04.2026" (dot-separated, D.M.Y)
      var m5 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
      if (m5) {
        var yr5 = parseInt(m5[3],10); if (yr5 < 100) yr5 += 2000;
        var dt5 = new Date(yr5, parseInt(m5[2],10)-1, parseInt(m5[1],10));
        if (!isNaN(dt5.getTime())) return dt5;
      }

      // Last resort: native parser (handles anything unusual we haven't explicitly covered)
      var d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }

// Translates raw, informally-typed root-cause text from transport staff into
// professional wording for anything leadership sees. Matches on distinctive
// substrings so minor wording/typo variants ("REFUSED DUE TO" vs "REFUSED TO
// ACCEPT DUE TO") merge into the same clean category instead of showing as
// separate duplicate rows.
function polishRootCause(raw) {
  var s = toStr(raw);
  if (!s) return s;
  var u = s.toUpperCase().trim();
  if (u.indexOf('MERCHANDISER') !== -1 && u.indexOf('ROUTE') !== -1) return 'Merchandiser Unavailable on Route';
  if (u.indexOf('FIRST GRV') !== -1 || u.indexOf('COLLECT GRV') !== -1) return 'Pending Goods Return Voucher Not Ready (GRV)';
  if (u.indexOf('SYSTEM NOT WORKING') !== -1) return 'Customer System Down';
  if (u.indexOf('SAME ITEM') !== -1 && u.indexOf('LPO') !== -1) return 'Duplicate Item Received Under Separate LPO';
  if (u.indexOf('HEAVY RAIN') !== -1 || (u.indexOf('ROAD CLOSURE') !== -1 && u.indexOf('RETURN') !== -1)) return 'Returned — Weather/Road Closure';
  if (u.indexOf('NO NEED STOCK') !== -1) return 'Declined — No Stock Requirement';
  if (u.indexOf('LPO DELETED') !== -1) return 'Declined — LPO Cancelled in Customer System';
  if (u.indexOf('RECEIVING CLOSED') !== -1) return 'Returned — Receiving Closed for the Day';
  if (u.indexOf('NO SPACE') !== -1) return 'Declined — Insufficient Storage Space';
  if (u.indexOf('NO SCHEDULE') !== -1) return 'Declined — No Delivery Schedule Confirmed';
  if (u.indexOf('PAYMENT') !== -1 && (u.indexOf('NOT READY') !== -1 || u.indexOf('NOT READ') !== -1)) return 'Returned — Customer Payment Not Ready (Cheque/Cash)';
  if (u.indexOf('LPO') !== -1 && u.indexOf('NOT') !== -1 && /L\w{0,3}ECTED/.test(u)) return 'Declined — LPO Not Reflected in Customer System';
  if (u.indexOf('SYSTEM UPDATING') !== -1) return 'Declined — Customer System Under Maintenance';
  if (u.indexOf('UNABLE TO ACCOMMODATE') !== -1) return 'Declined — Insufficient Storage Capacity Today';
  if (u.indexOf('RECEIVING TIME OVER') !== -1) return 'Declined — Outside Receiving Hours';
  // Fallback for anything not yet mapped: tidy spacing + Title Case, so it's
  // never worse than today even before someone adds a proper mapping for it.
  return s.replace(/\s+/g, ' ').trim().toLowerCase().replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

    var orgMap={}, monthMap={};
    var totalRej=0, totalDel=0, totalVal=0;
    var seenOrderVals={};

    for (var i=0; i<rows.length; i++) {
      var row=rows[i];
      var rej=isRej(row), del=isDel(row);
      if (!rej && !del) continue;
      var d=parseDate(row[RC.date]);
      var mo=d?d.getMonth()+1:null, day=d?d.getDate():null;
      var org=toStr(row[RC.org]).toUpperCase().replace('NON-FOOD','DGC');
      var root=polishRootCause(toStr(row[RC.root]));
      var cust=toStr(row[RC.cust]);
      var addr=RC.addr?toStr(row[RC.addr]):'';
      var area=toStr(row[RC.area]);
      var orderNo=RC.orderNo?toStr(row[RC.orderNo]):'';
      var rawVal=parseFloat(row[RC.value])||0;
      var val=(rej&&orderNo&&seenOrderVals[orderNo])?0:rawVal;
      if(rej&&orderNo&&!seenOrderVals[orderNo])seenOrderVals[orderNo]=rawVal;
      var typeStr=toStr(row[RC.type]||'').toUpperCase();
      var isFood=typeStr==='FOOD'||typeStr.startsWith('FOOD,');
      var isNF=typeStr.includes('NON FOOD')||typeStr.includes('NON-FOOD');
      var srcStr=toStr(row[RC.source]||'').toUpperCase();
      if (del) totalDel++;
      if (rej) { totalRej++; totalVal+=val; }
      if (org) {
        if (!orgMap[org]) orgMap[org]={tDel:0,tRej:0,val:0,food_rej:0,food_del:0,nonfood_rej:0,nonfood_del:0,ext_rej:0,ext_del:0,int_rej:0,int_del:0,food_val:0,nonfood_val:0,del:new Array(12).fill(0),rej:new Array(12).fill(0),reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{}};
        // Full per-org, per-month breakdown (food/nonfood/external/internal) — without this,
        // ORG + MONTH + SOURCE filters combined would fall back to all-org month totals,
        // which is what caused the >100% "Contribution to Rejection Rate" bug.
        if (mo) {
          if(!orgMap[org].byMonth) orgMap[org].byMonth={};
          if(!orgMap[org].byMonth[mo]) orgMap[org].byMonth[mo]={
            tRej:0,tDel:0,val:0,food_rej:0,food_del:0,nonfood_rej:0,nonfood_del:0,
            ext_rej:0,ext_del:0,int_rej:0,int_del:0,food_val:0,nonfood_val:0,
            reasons:{},custs:{},detail:{},
            food_reasons:{},food_custs:{},food_detail:{},
            nonfood_reasons:{},nonfood_custs:{},nonfood_detail:{},
            ext_reasons:{},ext_custs:{},ext_detail:{},
            int_reasons:{},int_custs:{},int_detail:{}
          };
        }
        var mb = mo ? orgMap[org].byMonth[mo] : null;
        if (del) {
          orgMap[org].tDel++; if(mo)orgMap[org].del[mo-1]++;
          if(isFood)orgMap[org].food_del++; else if(isNF)orgMap[org].nonfood_del++;
          if(srcStr==='EXTERNAL')orgMap[org].ext_del++; else if(srcStr==='INTERNAL')orgMap[org].int_del++;
          if(mb){
            mb.tDel++;
            if(isFood)mb.food_del++; else if(isNF)mb.nonfood_del++;
            if(srcStr==='EXTERNAL')mb.ext_del++; else if(srcStr==='INTERNAL')mb.int_del++;
          }
        }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val+=val; if(mo)orgMap[org].rej[mo-1]++;
          if(mb){
            mb.tRej++; mb.val+=val;
            if(root) mb.reasons[root]=(mb.reasons[root]||0)+1;
            if(cust) mb.custs[cust]=(mb.custs[cust]||0)+1;
            if(cust||root){
              var monthDetailKey = (cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
              mb.detail[monthDetailKey]=(mb.detail[monthDetailKey]||0)+1;
              if(isFood)  mb.food_detail[monthDetailKey]=(mb.food_detail[monthDetailKey]||0)+1;
              if(isNF)    mb.nonfood_detail[monthDetailKey]=(mb.nonfood_detail[monthDetailKey]||0)+1;
              if(srcStr==='EXTERNAL') mb.ext_detail[monthDetailKey]=(mb.ext_detail[monthDetailKey]||0)+1;
              if(srcStr==='INTERNAL') mb.int_detail[monthDetailKey]=(mb.int_detail[monthDetailKey]||0)+1;
            }
            if(isFood){ mb.food_rej++; mb.food_val+=val; if(root)mb.food_reasons[root]=(mb.food_reasons[root]||0)+1; if(cust)mb.food_custs[cust]=(mb.food_custs[cust]||0)+1; }
            else if(isNF){ mb.nonfood_rej++; mb.nonfood_val+=val; if(root)mb.nonfood_reasons[root]=(mb.nonfood_reasons[root]||0)+1; if(cust)mb.nonfood_custs[cust]=(mb.nonfood_custs[cust]||0)+1; }
            if(srcStr==='EXTERNAL'){ mb.ext_rej++; if(root)mb.ext_reasons[root]=(mb.ext_reasons[root]||0)+1; if(cust)mb.ext_custs[cust]=(mb.ext_custs[cust]||0)+1; }
            else if(srcStr==='INTERNAL'){ mb.int_rej++; if(root)mb.int_reasons[root]=(mb.int_reasons[root]||0)+1; if(cust)mb.int_custs[cust]=(mb.int_custs[cust]||0)+1; }
          }
          if(isFood){orgMap[org].food_rej++;orgMap[org].food_val+=val;}
          else if(isNF){orgMap[org].nonfood_rej++;orgMap[org].nonfood_val+=val;}
          if(srcStr==='EXTERNAL')orgMap[org].ext_rej++; else if(srcStr==='INTERNAL')orgMap[org].int_rej++;
          if(root)orgMap[org].reasons[root]=(orgMap[org].reasons[root]||0)+1;
          if(cust)orgMap[org].custs[cust]=(orgMap[org].custs[cust]||0)+1;
          if(area)orgMap[org].areas[area]=(orgMap[org].areas[area]||0)+1;
          // Detail: customer+address+rootcause combo
          if(cust||root){
            var detailKey = (cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
            orgMap[org].detail[detailKey]=(orgMap[org].detail[detailKey]||0)+1;
            if(isFood)  orgMap[org].food_detail[detailKey]=(orgMap[org].food_detail[detailKey]||0)+1;
            if(isNF)    orgMap[org].nonfood_detail[detailKey]=(orgMap[org].nonfood_detail[detailKey]||0)+1;
            if(srcStr==='EXTERNAL') orgMap[org].ext_detail[detailKey]=(orgMap[org].ext_detail[detailKey]||0)+1;
            if(srcStr==='INTERNAL') orgMap[org].int_detail[detailKey]=(orgMap[org].int_detail[detailKey]||0)+1;
          }
          // Per-type breakdown
          if(isFood){
            if(root)orgMap[org].food_reasons[root]=(orgMap[org].food_reasons[root]||0)+1;
            if(cust)orgMap[org].food_custs[cust]=(orgMap[org].food_custs[cust]||0)+1;
          } else if(isNF){
            if(root)orgMap[org].nonfood_reasons[root]=(orgMap[org].nonfood_reasons[root]||0)+1;
            if(cust)orgMap[org].nonfood_custs[cust]=(orgMap[org].nonfood_custs[cust]||0)+1;
          }
          // Per-source breakdown
          if(srcStr==='EXTERNAL'){
            if(root)orgMap[org].ext_reasons[root]=(orgMap[org].ext_reasons[root]||0)+1;
            if(cust)orgMap[org].ext_custs[cust]=(orgMap[org].ext_custs[cust]||0)+1;
          } else if(srcStr==='INTERNAL'){
            if(root)orgMap[org].int_reasons[root]=(orgMap[org].int_reasons[root]||0)+1;
            if(cust)orgMap[org].int_custs[cust]=(orgMap[org].int_custs[cust]||0)+1;
          }
        }
      }
      if (mo) {
        if (!monthMap[mo]) monthMap[mo]={days:{},tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},food_rej:0,nonfood_rej:0,ext_rej:0,int_rej:0,detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{},data:{}};
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++; monthMap[mo].val+=val;
          if(root)monthMap[mo].reasons[root]=(monthMap[mo].reasons[root]||0)+1;
          if(cust)monthMap[mo].custs[cust]=(monthMap[mo].custs[cust]||0)+1;
          if(area)monthMap[mo].areas[area]=(monthMap[mo].areas[area]||0)+1;
          if(day)monthMap[mo].days[day]=1;
          // Per-type breakdown for month
          if(isFood){
            monthMap[mo].food_rej++;
            if(root)monthMap[mo].food_reasons[root]=(monthMap[mo].food_reasons[root]||0)+1;
            if(cust)monthMap[mo].food_custs[cust]=(monthMap[mo].food_custs[cust]||0)+1;
          } else if(isNF){
            monthMap[mo].nonfood_rej++;
            if(root)monthMap[mo].nonfood_reasons[root]=(monthMap[mo].nonfood_reasons[root]||0)+1;
            if(cust)monthMap[mo].nonfood_custs[cust]=(monthMap[mo].nonfood_custs[cust]||0)+1;
          }
          // Per-source breakdown for month
          if(srcStr==='EXTERNAL'){
            monthMap[mo].ext_rej++;
            if(root)monthMap[mo].ext_reasons[root]=(monthMap[mo].ext_reasons[root]||0)+1;
            if(cust)monthMap[mo].ext_custs[cust]=(monthMap[mo].ext_custs[cust]||0)+1;
          } else if(srcStr==='INTERNAL'){
            monthMap[mo].int_rej++;
            if(root)monthMap[mo].int_reasons[root]=(monthMap[mo].int_reasons[root]||0)+1;
            if(cust)monthMap[mo].int_custs[cust]=(monthMap[mo].int_custs[cust]||0)+1;
          }
          // Month detail
          if(cust||root){
            var mdk=(cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
            monthMap[mo].detail[mdk]=(monthMap[mo].detail[mdk]||0)+1;
            if(isFood)  monthMap[mo].food_detail[mdk]=(monthMap[mo].food_detail[mdk]||0)+1;
            if(isNF)    monthMap[mo].nonfood_detail[mdk]=(monthMap[mo].nonfood_detail[mdk]||0)+1;
            if(srcStr==='EXTERNAL') monthMap[mo].ext_detail[mdk]=(monthMap[mo].ext_detail[mdk]||0)+1;
            if(srcStr==='INTERNAL') monthMap[mo].int_detail[mdk]=(monthMap[mo].int_detail[mdk]||0)+1;
          }
        }
        if (day) {
          if(!monthMap[mo].data[day])monthMap[mo].data[day]={tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},food_rej:0,nonfood_rej:0,ext_rej:0,int_rej:0,detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{}};
          if(del)monthMap[mo].data[day].tDel++;
          if(rej){
            monthMap[mo].data[day].tRej++; monthMap[mo].data[day].val+=val;
            if(root)monthMap[mo].data[day].reasons[root]=(monthMap[mo].data[day].reasons[root]||0)+1;
            if(cust)monthMap[mo].data[day].custs[cust]=(monthMap[mo].data[day].custs[cust]||0)+1;
            if(area)monthMap[mo].data[day].areas[area]=(monthMap[mo].data[day].areas[area]||0)+1;
            if(isFood){
              monthMap[mo].data[day].food_rej++;
              if(root)monthMap[mo].data[day].food_reasons[root]=(monthMap[mo].data[day].food_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].food_custs[cust]=(monthMap[mo].data[day].food_custs[cust]||0)+1;
            } else if(isNF){
              monthMap[mo].data[day].nonfood_rej++;
              if(root)monthMap[mo].data[day].nonfood_reasons[root]=(monthMap[mo].data[day].nonfood_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].nonfood_custs[cust]=(monthMap[mo].data[day].nonfood_custs[cust]||0)+1;
            }
            if(srcStr==='EXTERNAL'){
              monthMap[mo].data[day].ext_rej++;
              if(root)monthMap[mo].data[day].ext_reasons[root]=(monthMap[mo].data[day].ext_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].ext_custs[cust]=(monthMap[mo].data[day].ext_custs[cust]||0)+1;
            } else if(srcStr==='INTERNAL'){
              monthMap[mo].data[day].int_rej++;
              if(root)monthMap[mo].data[day].int_reasons[root]=(monthMap[mo].data[day].int_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].int_custs[cust]=(monthMap[mo].data[day].int_custs[cust]||0)+1;
            }
            // Day detail
            if(cust||root){
              var ddk=(cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
              monthMap[mo].data[day].detail[ddk]=(monthMap[mo].data[day].detail[ddk]||0)+1;
              if(isFood)  monthMap[mo].data[day].food_detail[ddk]=(monthMap[mo].data[day].food_detail[ddk]||0)+1;
              if(isNF)    monthMap[mo].data[day].nonfood_detail[ddk]=(monthMap[mo].data[day].nonfood_detail[ddk]||0)+1;
              if(srcStr==='EXTERNAL') monthMap[mo].data[day].ext_detail[ddk]=(monthMap[mo].data[day].ext_detail[ddk]||0)+1;
              if(srcStr==='INTERNAL') monthMap[mo].data[day].int_detail[ddk]=(monthMap[mo].data[day].int_detail[ddk]||0)+1;
            }
          }
        }
      }
    }

    function fmtVal(v){return v>=1000000?'AED '+(v/1000000).toFixed(2)+'M':'AED '+Math.round(v/1000)+'K';}
    function top10(obj){return Object.keys(obj).map(function(l){return{l:l,n:obj[l]};}).sort(function(a,b){return b.n-a.n;}).slice(0,10);}
    function top8c(obj){return Object.keys(obj).map(function(n){return{n:n,c:obj[n],v:''};}).sort(function(a,b){return b.c-a.c;}).slice(0,20);}
    function top6a(obj){return Object.keys(obj).map(function(a){return{a:a,n:obj[a]};}).sort(function(a,b){return b.n-a.n;}).slice(0,6);}
    function topDetail(obj){return Object.keys(obj).map(function(k){var p=k.split('|||');return{cust:p[0],addr:p[1],root:p[2],n:obj[k]};}).sort(function(a,b){return b.n-a.n;}).slice(0,100);}

    var allR={},allC={},allA={},allDetail={},allFoodDetail={},allNFDetail={},allExtDetail={},allIntDetail={},allDel=new Array(12).fill(0),allRej=new Array(12).fill(0);
    var allFoodRej=0,allNFRej=0,allExtRej=0,allIntRej=0,allFoodDel=0,allNFDel=0,allFoodVal=0,allNFVal=0;
    var allFoodR={},allFoodC={},allNFR={},allNFC={},allExtR={},allExtC={},allIntR={},allIntC={};
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      Object.keys(v.reasons).forEach(function(k){allR[k]=(allR[k]||0)+v.reasons[k];});
      Object.keys(v.custs).forEach(function(k){allC[k]=(allC[k]||0)+v.custs[k];});
      Object.keys(v.areas).forEach(function(k){allA[k]=(allA[k]||0)+v.areas[k];});
      Object.keys(v.detail||{}).forEach(function(k){allDetail[k]=(allDetail[k]||0)+v.detail[k];});
      Object.keys(v.food_detail||{}).forEach(function(k){allFoodDetail[k]=(allFoodDetail[k]||0)+v.food_detail[k];});
      Object.keys(v.nonfood_detail||{}).forEach(function(k){allNFDetail[k]=(allNFDetail[k]||0)+v.nonfood_detail[k];});
      Object.keys(v.ext_detail||{}).forEach(function(k){allExtDetail[k]=(allExtDetail[k]||0)+v.ext_detail[k];});
      Object.keys(v.int_detail||{}).forEach(function(k){allIntDetail[k]=(allIntDetail[k]||0)+v.int_detail[k];});
      Object.keys(v.food_reasons||{}).forEach(function(k){allFoodR[k]=(allFoodR[k]||0)+v.food_reasons[k];});
      Object.keys(v.food_custs||{}).forEach(function(k){allFoodC[k]=(allFoodC[k]||0)+v.food_custs[k];});
      Object.keys(v.nonfood_reasons||{}).forEach(function(k){allNFR[k]=(allNFR[k]||0)+v.nonfood_reasons[k];});
      Object.keys(v.nonfood_custs||{}).forEach(function(k){allNFC[k]=(allNFC[k]||0)+v.nonfood_custs[k];});
      Object.keys(v.ext_reasons||{}).forEach(function(k){allExtR[k]=(allExtR[k]||0)+v.ext_reasons[k];});
      Object.keys(v.ext_custs||{}).forEach(function(k){allExtC[k]=(allExtC[k]||0)+v.ext_custs[k];});
      Object.keys(v.int_reasons||{}).forEach(function(k){allIntR[k]=(allIntR[k]||0)+v.int_reasons[k];});
      Object.keys(v.int_custs||{}).forEach(function(k){allIntC[k]=(allIntC[k]||0)+v.int_custs[k];});
      v.del.forEach(function(d,i){allDel[i]+=d;}); v.rej.forEach(function(r,i){allRej[i]+=r;});
      allFoodRej+=(v.food_rej||0); allNFRej+=(v.nonfood_rej||0);
      allExtRej+=(v.ext_rej||0); allIntRej+=(v.int_rej||0);
      allFoodDel+=(v.food_del||0); allNFDel+=(v.nonfood_del||0);
      allFoodVal+=(v.food_val||0); allNFVal+=(v.nonfood_val||0);
    });

    var monthsOut={};
    Object.keys(monthMap).forEach(function(mo){
      var md=monthMap[mo]; var dataOut={};
      Object.keys(md.data).forEach(function(day){
        var dd=md.data[day];
        dataOut[day]={tDel:dd.tDel,tRej:dd.tRej,val:fmtVal(dd.val),reasons:top10(dd.reasons),custs:top8c(dd.custs),areas:top6a(dd.areas),food_rej:dd.food_rej||0,nonfood_rej:dd.nonfood_rej||0,ext_rej:dd.ext_rej||0,int_rej:dd.int_rej||0,food_reasons:top10(dd.food_reasons||{}),food_custs:top8c(dd.food_custs||{}),nonfood_reasons:top10(dd.nonfood_reasons||{}),nonfood_custs:top8c(dd.nonfood_custs||{}),ext_reasons:top10(dd.ext_reasons||{}),ext_custs:top8c(dd.ext_custs||{}),int_reasons:top10(dd.int_reasons||{}),int_custs:top8c(dd.int_custs||{}),detail:topDetail(dd.detail||{}),food_detail:topDetail(dd.food_detail||{}),nonfood_detail:topDetail(dd.nonfood_detail||{}),ext_detail:topDetail(dd.ext_detail||{}),int_detail:topDetail(dd.int_detail||{})};
      });
      monthsOut[mo]={days:Object.keys(md.days).map(Number).sort(function(a,b){return a-b;}),tDel:md.tDel,tRej:md.tRej,val:fmtVal(md.val),food_rej:md.food_rej||0,nonfood_rej:md.nonfood_rej||0,ext_rej:md.ext_rej||0,int_rej:md.int_rej||0,reasons:top10(md.reasons),custs:top8c(md.custs||{}),areas:top6a(md.areas||{}),food_reasons:top10(md.food_reasons||{}),food_custs:top8c(md.food_custs||{}),nonfood_reasons:top10(md.nonfood_reasons||{}),nonfood_custs:top8c(md.nonfood_custs||{}),ext_reasons:top10(md.ext_reasons||{}),ext_custs:top8c(md.ext_custs||{}),int_reasons:top10(md.int_reasons||{}),int_custs:top8c(md.int_custs||{}),detail:topDetail(md.detail||{}),food_detail:topDetail(md.food_detail||{}),nonfood_detail:topDetail(md.nonfood_detail||{}),ext_detail:topDetail(md.ext_detail||{}),int_detail:topDetail(md.int_detail||{}),data:dataOut};
    });

    var orgsOut={all:{tDel:totalDel,tRej:totalRej,val:fmtVal(totalVal),food_rej:allFoodRej,food_del:allFoodDel,nonfood_rej:allNFRej,nonfood_del:allNFDel,ext_rej:allExtRej,int_rej:allIntRej,food_val:fmtVal(allFoodVal),nonfood_val:fmtVal(allNFVal),del:allDel,rej:allRej,reasons:top10(allR),custs:top8c(allC),areas:top6a(allA),detail:topDetail(allDetail),food_detail:topDetail(allFoodDetail||{}),nonfood_detail:topDetail(allNFDetail||{}),ext_detail:topDetail(allExtDetail||{}),int_detail:topDetail(allIntDetail||{}),food_reasons:top10(allFoodR),food_custs:top8c(allFoodC),nonfood_reasons:top10(allNFR),nonfood_custs:top8c(allNFC),ext_reasons:top10(allExtR),ext_custs:top8c(allExtC),int_reasons:top10(allIntR),int_custs:top8c(allIntC)}};
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      var byMonthOut = {};
      if (v.byMonth) {
        Object.keys(v.byMonth).forEach(function(mo){
          var mData = v.byMonth[mo];
          byMonthOut[mo] = {
            tRej: mData.tRej||0, tDel: mData.tDel||0, val: fmtVal(mData.val||0),
            food_rej: mData.food_rej||0, food_del: mData.food_del||0, food_val: fmtVal(mData.food_val||0),
            nonfood_rej: mData.nonfood_rej||0, nonfood_del: mData.nonfood_del||0, nonfood_val: fmtVal(mData.nonfood_val||0),
            ext_rej: mData.ext_rej||0, ext_del: mData.ext_del||0,
            int_rej: mData.int_rej||0, int_del: mData.int_del||0,
            reasons: top10(mData.reasons||{}), custs: top8c(mData.custs||{}), detail: topDetail(mData.detail||{}),
            food_reasons: top10(mData.food_reasons||{}), food_custs: top8c(mData.food_custs||{}), food_detail: topDetail(mData.food_detail||{}),
            nonfood_reasons: top10(mData.nonfood_reasons||{}), nonfood_custs: top8c(mData.nonfood_custs||{}), nonfood_detail: topDetail(mData.nonfood_detail||{}),
            ext_reasons: top10(mData.ext_reasons||{}), ext_custs: top8c(mData.ext_custs||{}), ext_detail: topDetail(mData.ext_detail||{}),
            int_reasons: top10(mData.int_reasons||{}), int_custs: top8c(mData.int_custs||{}), int_detail: topDetail(mData.int_detail||{})
          };
        });
      }
      orgsOut[org]={tDel:v.tDel,tRej:v.tRej,val:fmtVal(v.val),food_rej:v.food_rej||0,food_del:v.food_del||0,nonfood_rej:v.nonfood_rej||0,nonfood_del:v.nonfood_del||0,ext_rej:v.ext_rej||0,int_rej:v.int_rej||0,food_val:fmtVal(v.food_val||0),nonfood_val:fmtVal(v.nonfood_val||0),del:v.del,rej:v.rej,reasons:top10(v.reasons),custs:top8c(v.custs),areas:top6a(v.areas),detail:topDetail(v.detail||{}),food_detail:topDetail(v.food_detail||{}),nonfood_detail:topDetail(v.nonfood_detail||{}),ext_detail:topDetail(v.ext_detail||{}),int_detail:topDetail(v.int_detail||{}),food_reasons:top10(v.food_reasons||{}),food_custs:top8c(v.food_custs||{}),nonfood_reasons:top10(v.nonfood_reasons||{}),nonfood_custs:top8c(v.nonfood_custs||{}),ext_reasons:top10(v.ext_reasons||{}),ext_custs:top8c(v.ext_custs||{}),int_reasons:top10(v.int_reasons||{}),int_custs:top8c(v.int_custs||{}),byMonth:byMonthOut};
    });

    rejectionData={uploadedAt:new Date().toISOString(),uploadedBy:req.body.uploadedBy||'Admin',fileName:req.file.originalname,totalOrders:totalRej+totalDel,orgs:orgsOut,months:monthsOut};

    // Save to DB + file
    var dbOk = await dbSaveRejection(rejectionData.uploadedBy, rejectionData.fileName, rejectionData.totalOrders, orgsOut, monthsOut);
    saveJSON(REJECTION_FILE, rejectionData);
    console.log('Rejection saved:', totalRej, 'rej', totalDel, 'del', dbOk ? '(DB+file)' : '(file only)');

    res.json({ success:true, summary:{totalRej:totalRej,totalDel:totalDel,fileName:req.file.originalname} });
  } catch(e) {
    console.error('Rejection upload error:', e.message, e.stack);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.get('/api/rejection/status', function(req, res) {
  res.setHeader('Content-Type', 'application/json');
  if (!rejectionData) return res.json({ hasData:false });
  res.json({ hasData:true, uploadedAt:rejectionData.uploadedAt, uploadedBy:rejectionData.uploadedBy, fileName:rejectionData.fileName, totalOrders:rejectionData.totalOrders, orgs:rejectionData.orgs, months:rejectionData.months, needsReupload:rejectionData.needsReupload||false });
});

// ── REJECTION EXCEL EXPORT (server-side, 2 sheets: Executive Summary + Detail, styled) ──
app.post('/api/rejection/export-excel', async function(req, res) {
  try {
    var body = req.body || {};
    var summary = body.summary || {};
    var detailRows = body.detailRows || [];
    var filterStr = body.filterStr || 'All Filters';

    var totalRej = summary.tRej || 0;
    var totalDel = summary.tDel || 0;
    var rate = (totalRej + totalDel) > 0 ? (totalRej/(totalRej+totalDel)*100).toFixed(2)+'%' : '0%';
    var topReasons = (summary.reasons || []).slice(0, 5);
    var topCusts = (summary.custs || []).slice(0, 5);

    // For each top root cause, find the single customer+branch driving the most cases of it
    var rootCauseTop = {};
    detailRows.forEach(function(r){
      var key = r.root || 'Unknown';
      if(!rootCauseTop[key] || (r.n||0) > rootCauseTop[key].n){
        rootCauseTop[key] = { cust: r.cust||'—', addr: r.addr||'—', n: r.n||0 };
      }
    });

    // Group detail rows by customer to find repeat offenders (same customer, multiple branch addresses)
    // Track rejection count PER branch (not just presence) so we can list the actual branches
    var custGroups = {};
    detailRows.forEach(function(r){
      var key = r.cust || 'Unknown';
      if(!custGroups[key]) custGroups[key] = { branches: {}, total: 0 };
      var addrKey = r.addr || 'Unknown address';
      custGroups[key].branches[addrKey] = (custGroups[key].branches[addrKey] || 0) + (r.n || 0);
      custGroups[key].total += (r.n || 0);
    });
    var repeatOffenders = Object.keys(custGroups)
      .map(function(name){
        var branchList = Object.keys(custGroups[name].branches)
          .map(function(a){ return { addr:a, n:custGroups[name].branches[a] }; })
          .sort(function(x,y){ return y.n - x.n; });
        return { name:name, branchCount: branchList.length, total: custGroups[name].total, branches: branchList };
      })
      .filter(function(c){ return c.branchCount >= 2; })
      .sort(function(a,b){ return (b.branchCount - a.branchCount) || (b.total - a.total); })
      .slice(0, 8);

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', WHITE = 'FFFFFFFF';
    function styleHeaderRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
        cell.font = { bold:true, color:{argb:GOLD}, size:11 };
        cell.alignment = { vertical:'middle' };
      });
    }
    function styleSectionRow(row){
      row.font = { bold:true, color:{argb:DARKBG}, size:12 };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:GOLD} }; });
    }
    function styleTotalRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:LIGHTGOLD} };
        cell.font = { bold:true, color:{argb:DARKBG} };
      });
    }
    function styleCustRow(row){
      row.font = { bold:true };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFEFEFEF'} }; });
    }

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    // ---- Sheet 1: Executive Summary ----
    var es = wb.addWorksheet('Executive Summary');
    es.columns = [{width:38},{width:32},{width:14},{width:14},{width:34}];

    var titleRow = es.addRow(['AKI GROUP — REJECTION EXECUTIVE SUMMARY']);
    titleRow.font = { bold:true, size:15, color:{argb:GOLD} };
    es.mergeCells('A1:E1');
    es.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
    es.addRow(['Generated', new Date().toLocaleString('en-AE')]);
    es.addRow(['Filters Applied', filterStr]);
    es.addRow([]);
    styleSectionRow(es.addRow(['KEY METRICS']));
    es.addRow(['Total Rejections', totalRej]);
    es.addRow(['Total Delivered', totalDel]);
    es.addRow(['Rejection Rate', rate]);
    es.addRow(['Value at Risk', summary.val || '—']);
    es.addRow([]);

    // TOP 5 ROOT CAUSES — now shows WHO (customer + branch) is the biggest driver of each cause
    styleHeaderRow(es.addRow(['ROOT CAUSE', 'TOP CUSTOMER DRIVING IT', 'COUNT', '% OF TOTAL', 'TOP BRANCH (CASES)']));
    topReasons.forEach(function(r){
      var top = rootCauseTop[r.l] || { cust:'—', addr:'—', n:0 };
      es.addRow([r.l, top.cust, r.n, totalRej>0 ? ((r.n/totalRej*100).toFixed(1)+'%') : '0%', top.addr + ' ('+top.n+')']);
    });
    es.addRow([]);

    styleHeaderRow(es.addRow(['TOP 5 CUSTOMERS BY REJECTIONS', '', 'COUNT', '% OF TOTAL', '']));
    topCusts.forEach(function(c){
      es.addRow([c.n, '', c.c, totalRej>0 ? ((c.c/totalRej*100).toFixed(1)+'%') : '0%', '']);
    });
    es.addRow([]);

    // Repeat Offenders — customer name AND the actual branch addresses driving their total
    styleHeaderRow(es.addRow(['REPEAT OFFENDERS — CUSTOMER / BRANCH', '', 'BRANCHES', 'TOTAL REJECTIONS', '']));
    if(repeatOffenders.length){
      repeatOffenders.forEach(function(o){
        styleCustRow(es.addRow([o.name, '', o.branchCount, o.total, '']));
        o.branches.slice(0, 6).forEach(function(b){
          es.addRow(['     • '+b.addr, '', '', b.n, '']);
        });
      });
    } else {
      es.addRow(['No customer repeats across multiple branches in this filtered view.']);
    }
    es.addRow([]);

    styleSectionRow(es.addRow(['RECOMMENDED ACTIONS']));
    if (topReasons[0]) {
      var t0 = rootCauseTop[topReasons[0].l] || {};
      es.addRow(['1. Prioritize a fix for "'+topReasons[0].l+'" — top root cause at '+topReasons[0].n+' rejections ('+(totalRej>0?(topReasons[0].n/totalRej*100).toFixed(1):'0')+'% of total). Biggest driver: "'+(t0.cust||'—')+'" at '+(t0.addr||'—')+' ('+(t0.n||0)+' cases).']);
    }
    if (topReasons[1]) es.addRow(['2. Address "'+topReasons[1].l+'" next — '+topReasons[1].n+' rejections.']);
    if (topCusts[0]) es.addRow(['3. Engage account owner for "'+topCusts[0].n+'" — highest-rejecting customer with '+topCusts[0].c+' cases.']);
    if (repeatOffenders[0]) es.addRow(['4. Investigate "'+repeatOffenders[0].name+'" across '+repeatOffenders[0].branchCount+' branches (see list above) — recurring issue, not a one-off location problem.']);
    es.addRow(['5. Re-check merchandiser/route scheduling if route-related causes dominate the list above.']);

    // ---- Sheet 2: Rejection Detail ----
    var det = wb.addWorksheet('Rejection Detail');
    det.columns = [{width:5},{width:32},{width:45},{width:30},{width:14},{width:12}];
    styleHeaderRow(det.addRow(['#','Customer Name','Full Address','Final Root Cause','Rejection Count','% of Total']));
    var sumN = 0;
    detailRows.forEach(function(r, i){
      var pct = totalRej>0 ? ((r.n/totalRej*100).toFixed(2)+'%') : '0%';
      sumN += (r.n||0);
      det.addRow([i+1, r.cust||'', r.addr||'', r.root||'', r.n||0, pct]);
    });
    var pctTotal = totalRej>0 ? ((sumN/totalRej*100).toFixed(2)+'%') : '0%';
    styleTotalRow(det.addRow(['', 'TOTAL', '', '', sumN, pctTotal]));

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Rejection_Report_'+Date.now()+'.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch(e) {
    console.error('rejection export-excel error:', e.message);
    res.status(500).json({ error: 'Export failed: '+e.message });
  }
});

// ── DISPATCH DROP ANALYSIS + ROUTE SUMMARY EXCEL EXPORT ──
app.post('/api/dispatch/export-excel', async function(req, res) {
  try {
    var body = req.body || {};
    var ca = body.cost_analysis || {};
    var topRoutes = body.top_routes || [];
    var repeatLocs = body.repeat_locations || [];
    var cityTypeCross = body.city_type_cross || {};

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', REQBLUE = 'FFD6E8FF', AVOIDRED = 'FFFDE0DE';
    function styleHeaderRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
        cell.font = { bold:true, color:{argb:GOLD}, size:11 };
        cell.alignment = { vertical:'middle' };
      });
    }
    function styleSectionRow(row){
      row.font = { bold:true, color:{argb:DARKBG}, size:12 };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:GOLD} }; });
    }
    function styleTotalRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:LIGHTGOLD} };
        cell.font = { bold:true, color:{argb:DARKBG} };
      });
    }

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    var avoidableCount = ca.repeat_location_avoidable_count || 0;
    var totalRepeatCount = ca.repeat_location_count || 0;
    var avoidableRows = repeatLocs.filter(function(l){ return !l.is_legitimate_split; });
    var requiredRows = repeatLocs.filter(function(l){ return l.is_legitimate_split; });

    // ---- Sheet 1: Executive Summary ----
    var es = wb.addWorksheet('Executive Summary');
    es.columns = [{width:36},{width:22},{width:16},{width:16},{width:16}];

    var titleRow = es.addRow(['AKI GROUP — DAILY DISPATCH DROP ANALYSIS']);
    titleRow.font = { bold:true, size:15, color:{argb:GOLD} };
    es.mergeCells('A1:E1');
    es.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
    es.addRow(['Date', body.date || new Date().toLocaleDateString('en-AE')]);
    es.addRow(['Generated', new Date().toLocaleString('en-AE')]);
    es.addRow([]);

    styleSectionRow(es.addRow(['KEY METRICS']));
    es.addRow(['Total Orders', body.total_orders || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Value (AED)', Math.round(body.total_value || 0)]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Routes', body.total_routes || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Drivers', body.total_drivers || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Drops', body.total_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Own Fleet Drops', ca.own_fleet_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['3PL Drops (billed separately)', ca.pl_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow([]);

    styleSectionRow(es.addRow(['REPEAT-VISIT ADDRESSES — ACTION SUMMARY']));
    es.addRow(['Total addresses visited by 2+ routes today', totalRepeatCount]);
    es.addRow(['— Required (different order types, separate trucks needed)', requiredRows.length]);
    es.addRow(['— Avoidable (same order type, worth questioning)', avoidableCount]);
    es.addRow([]);

    var RED = 'FFB0201A';
    var exceptionRows = avoidableRows.filter(function(l){ return l.is_high_value_exception; });
    var genuineAvoidableRows = avoidableRows.filter(function(l){ return !l.is_high_value_exception; });

    if (exceptionRows.length) {
      var exceptionsByValue = exceptionRows.slice().sort(function(a, b) { return (b.total_value || 0) - (a.total_value || 0); });
      styleHeaderRow(es.addRow(['⚠ HIGH-VALUE EXCEPTIONS (>AED 100,000)', 'CUSTOMER', 'ROUTES (TYPE · ORDERS · VALUE)', 'TOTAL VALUE (AED)', '']));
      es.addRow(['These are same-order-type duplicates, but the value is large enough that they may be a genuinely large order needing a capacity split — verify before treating as a routing mistake.']);
      exceptionsByValue.forEach(function(l){
        var routesLabel = (l.route_types||[]).map(function(rt){ return rt.route+' ('+rt.types.join('+')+', '+(rt.order_count||0)+' orders, AED '+(rt.value||0).toLocaleString()+')'; }).join(', ') || (l.routes||[]).join(', ');
        var row = es.addRow([l.location_id, l.customer, routesLabel, l.total_value || 0, '']);
        row.getCell(4).numFmt = '#,##0';
        row.font = { color:{argb:'FF8B6914'} };
      });
      es.addRow([]);
    }

    if (genuineAvoidableRows.length) {
      styleHeaderRow(es.addRow(['TOP AVOIDABLE REPEAT DROPS', 'CUSTOMER', 'ROUTES (TYPE · ORDERS · VALUE)', 'TOTAL VALUE (AED)', '']));
      genuineAvoidableRows.slice(0, 10).forEach(function(l){
        var routesLabel = (l.route_types||[]).map(function(rt){ return rt.route+' ('+rt.types.join('+')+', '+(rt.order_count||0)+' orders, AED '+(rt.value||0).toLocaleString()+')'; }).join(', ') || (l.routes||[]).join(', ');
        var row = es.addRow([l.location_id, l.customer, routesLabel, l.total_value || 0, '']);
        row.getCell(4).numFmt = '#,##0';
        row.font = { color:{argb:RED} };
      });
      es.addRow([]);
    }

    styleSectionRow(es.addRow(['RECOMMENDED ACTIONS']));
    if (genuineAvoidableRows.length) {
      var biggest = genuineAvoidableRows[0];
      es.addRow(['1. Raise the ' + genuineAvoidableRows.length + ' avoidable repeat-drop addresses with the transport team — same order type sent on 2+ separate trucks to the same address.']);
      es.addRow(['2. Start with "' + biggest.customer + '" (Location ' + biggest.location_id + ') — highest combined value at AED ' + (biggest.total_value||0).toLocaleString() + ' split across routes ' + (biggest.routes||[]).join(', ') + '.']);
      es.addRow(['3. See "Repeat Location Detail" tab for the full list with per-route order values — check whether each split was due to genuine order size before assuming it was a routing error.']);
    } else {
      es.addRow(['No avoidable repeat drops found today — all multi-route visits were legitimate Food/Non-Food/3PL splits.']);
    }
    if (exceptionRows.length) {
      es.addRow(['4. ' + exceptionRows.length + ' high-value exception(s) flagged above (>AED 100,000) — review order size before raising these with the transport team, as a large order may genuinely require 2 trucks.']);
    }

    // ---- Sheet 2: Route Summary ----
    var rt = wb.addWorksheet('Route Summary');
    rt.columns = [{width:14},{width:30},{width:14},{width:12},{width:12},{width:14},{width:16}];
    styleHeaderRow(rt.addRow(['Route', 'Type', 'Vehicle', 'Orders', 'Drivers', 'Locations (Drops)', 'Value (AED)']));
    var sumOrders = 0, sumDrops = 0, sumValue = 0;
    topRoutes.forEach(function(r){
      sumOrders += r.orders || 0; sumDrops += r.drops || 0; sumValue += r.value || 0;
      var typeLabel = (r.types || []).join(', ');
      var vehicleLabel = r.isPartitionVehicle ? 'Partition' : 'Single-Type';
      var row = rt.addRow([r.route, typeLabel, vehicleLabel, r.orders || 0, r.driverCount || 0, r.drops || 0, Math.round(r.value || 0)]);
      row.getCell(7).numFmt = '#,##0';
      if (r.isPartitionVehicle) { row.getCell(3).font = { bold:true, color:{argb:GOLD} }; }
    });
    var totalRow = rt.addRow(['TOTAL', '', '', sumOrders, '', sumDrops, Math.round(sumValue)]);
    totalRow.getCell(7).numFmt = '#,##0';
    styleTotalRow(totalRow);

    // ---- Sheet 3: Transport Cost (rate-card based, only if truck-type column present) ----
    if (body.truck_cost_estimate && body.truck_cost_estimate.available) {
      var tce = body.truck_cost_estimate;
      var tc = wb.addWorksheet('Transport Cost');
      tc.columns = [{width:26},{width:14},{width:14},{width:16}];
      styleSectionRow(tc.addRow(['SUMMARY']));
      tc.addRow(['Total Estimated Cost (AED)', tce.total_estimated_cost || 0]).getCell(2).numFmt = '#,##0';
      tc.addRow(['Total Drops Billed', tce.total_drops_billed || 0]);
      tc.addRow(['Total Vehicles Used', tce.total_vehicles || 0]);
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY VEHICLE TYPE', 'Rate (AED/Drop)', 'Drops Billed', 'Est. Cost (AED)']));
      (tce.by_type || []).forEach(function(t){
        var row = tc.addRow([t.label, t.rate, t.drop_count, t.estimated_cost]);
        row.getCell(4).numFmt = '#,##0';
      });
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY REGION', '', 'Drops', 'Est. Cost (AED)']));
      (tce.by_region || []).forEach(function(r){
        var row = tc.addRow([r.city, '', r.total_drops, r.total_cost]);
        row.getCell(4).numFmt = '#,##0';
      });
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY FOOD / NON-FOOD', '', 'Drops', 'Est. Cost (AED)']));
      (tce.by_food_type || []).forEach(function(f){
        var row = tc.addRow([f.category, '', f.drop_count, f.estimated_cost]);
        row.getCell(4).numFmt = '#,##0';
      });

      var unmatchedKeysExport = Object.keys(tce.unmatched_truck_types || {});
      if (unmatchedKeysExport.length) {
        tc.addRow([]);
        var warnRow = tc.addRow(['⚠ Unmatched truck-type text (excluded from cost, not guessed):']);
        warnRow.font = { italic: true, color: {argb:'FFB0201A'} };
        unmatchedKeysExport.forEach(function(k){
          tc.addRow(['  ' + k, '', tce.unmatched_truck_types[k], '']);
        });
      }
    }

    // ---- Sheet 4: In-House vs Hired Drivers ----
    if (body.driver_source_split) {
      var dss = body.driver_source_split;
      var dh = wb.addWorksheet('In-House vs Hired Drivers');
      dh.columns = [{width:30},{width:32},{width:16},{width:12},{width:12},{width:16}];
      styleSectionRow(dh.addRow(['SUMMARY']));
      dh.addRow(['In-House Drivers', dss.inhouse.driver_count || 0, '', dss.inhouse.orders||0, dss.inhouse.drops||0, dss.inhouse.value||0]).getCell(6).numFmt = '#,##0';
      var hiredSummaryRow = dh.addRow(['Hired Drivers (no name on file)', dss.hired.driver_count || 0, '', dss.hired.orders||0, dss.hired.drops||0, dss.hired.value||0]);
      hiredSummaryRow.getCell(6).numFmt = '#,##0';
      hiredSummaryRow.font = { color:{argb:'FFB0201A'} };
      dh.addRow([]);
      styleHeaderRow(dh.addRow(['Phone/ID', 'Customer', 'Type', 'Orders', 'Drops', 'Value (AED)']));
      (dss.hired_driver_details || []).forEach(function(d){
        var custLabel = (d.customers && d.customers.length) ? d.customers.join(', ') : '—';
        var typeLabel2 = (d.types && d.types.length) ? d.types.join(', ') : '—';
        var row2 = dh.addRow([d.name, custLabel, typeLabel2, d.orders || 0, d.drops || 0, d.value || 0]);
        row2.getCell(6).numFmt = '#,##0';
        row2.getCell(1).font = { color:{argb:'FFB0201A'} };
      });
    }

    // ---- Sheet 5: Repeat Location Detail (one row per individual order for full traceability) ----
    var rl = wb.addWorksheet('Repeat Location Detail');
    rl.columns = [{width:14},{width:30},{width:40},{width:12},{width:26},{width:10},{width:14},{width:14},{width:16}];
    styleHeaderRow(rl.addRow(['Location ID', 'Customer', 'Address', 'Route', 'Order Code', 'Type', 'Order Value (AED)', 'Status', 'Location Total (AED)']));
    var rlRowCount = 1;
    repeatLocs.forEach(function(l){
      var statusText = l.is_legitimate_split ? 'Required' : (l.is_high_value_exception ? 'Exception — Review' : 'Avoidable');
      var statusColor = l.is_legitimate_split ? REQBLUE : (l.is_high_value_exception ? 'FFFFE9A8' : AVOIDRED);
      var fontColor = l.is_legitimate_split ? 'FF1B5E9E' : (l.is_high_value_exception ? 'FF8B6914' : 'FFB0201A');
      (l.route_types || []).forEach(function(rt2){
        var orders = (rt2.orders && rt2.orders.length) ? rt2.orders : [{ order_code:'', type:(rt2.types||[])[0]||'', value: rt2.value }];
        orders.forEach(function(ord){
          var row = rl.addRow([
            l.location_id,
            l.customer,
            l.address,
            rt2.route,
            ord.order_code || '—',
            ord.type || '',
            ord.value || 0,
            statusText,
            l.total_value || 0
          ]);
          rlRowCount++;
          row.getCell(7).numFmt = '#,##0';
          row.getCell(9).numFmt = '#,##0';
          row.getCell(8).fill = { type:'pattern', pattern:'solid', fgColor:{argb: statusColor} };
          row.getCell(8).font = { bold:true, color:{argb: fontColor} };
          row.getCell(9).font = { bold:true, color:{argb: fontColor} };
        });
      });
    });
    // Enable AutoFilter on the header row so the Status column (blue=Required, red=Avoidable,
    // amber=Exception) can be filtered directly in Excel — every row now carries its own
    // Status value, so filtering shows complete, self-contained rows, not blank gaps.
    rl.autoFilter = { from: { row: 1, column: 1 }, to: { row: rlRowCount, column: 9 } };

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Dispatch_Drop_Analysis_'+Date.now()+'.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch(e) {
    console.error('dispatch export-excel error:', e.message);
    res.status(500).json({ error: 'Export failed: '+e.message });
  }
});

// ── SHARED PASSWORD CHECK FOR UPLOAD STARS ──
app.post('/api/backlog/check-password', requireAuth, async function(req, res) {
  try {
    var { password } = req.body;
    if (!password) return res.json({ ok: false });
    // Check against the logged-in user's own password
    var result = await pool.query('SELECT password_hash FROM users WHERE id=$1', [req.user.uid]);
    if (!result.rows[0]) return res.json({ ok: false });
    var match = await bcrypt.compare(password, result.rows[0].password_hash);
    res.json({ ok: match });
  } catch(e) { res.status(500).json({ ok: false, error: e.message }); }
});

// Caption must always have full DB access, regardless of what's loaded in the requesting browser's
// current session/tab. Pull directly from the server-side DB-backed globals every time.
function buildServerDataContext() {
  var parts = [];
  function addJSON(obj, label, maxLen) {
    try {
      if (obj === null || obj === undefined) { parts.push('=== ' + label + ': no file uploaded ==='); return; }
      var s = JSON.stringify(obj);
      if (s && s.length > maxLen) s = s.substring(0, maxLen) + '...(truncated)';
      parts.push('=== ' + label + ' (live from database) ===\n' + s);
    } catch(e) { parts.push('=== ' + label + ': error reading data ==='); }
  }

  addJSON(rejectionData ? { orgs: rejectionData.orgs, months: rejectionData.months } : null, 'REJECTION_DB', 4500);
  addJSON(currentDispatch ? { date: currentDispatch.date, summary: currentDispatch.summary } : null, 'DISPATCH_DB_LATEST', 3000);
  addJSON(Object.keys(dispatchHistory || {}).sort().reverse().slice(0, 14), 'DISPATCH_AVAILABLE_DATES', 500);
  addJSON(salesData ? salesData.summary : null, 'ORDER_BOOKING_DB (Todays Order Booking dashboard — BOOKED orders, NOT dispatch)', 3000);
  addJSON(returnsData ? returnsData.summary : null, 'RETURNS_DB', 3000);
  addJSON(backlogData ? backlogData.summary : null, 'BACKLOG_DB (WH Backlog)', 3000);
  addJSON(automationData ? { summary: automationData.summary, latestMonth: automationData.latestMonth, latestRate: automationData.latestRate, sortedMonths: automationData.sortedMonths } : null, 'AUTOMATION_DB', 2500);
  addJSON(deliveryData ? deliveryData.summary : null, 'DELIVERY_DB', 2000);
  addJSON(genInfoData ? genInfoData.rows : null, 'TEAM_DB', 3000);

  return parts.join('\n\n');
}

app.post('/api/voice', requireAuth, async function(req, res) {
  try {
    var text = req.body.text || '';
    var clientContext = req.body.context || '';
    var context = clientContext;
    try {
      context = clientContext + '\n\n' + buildServerDataContext();
    } catch(ctxErr) {
      console.error('buildServerDataContext failed, falling back to client-only context:', ctxErr.message);
    }
    var tab = req.body.tab || 'dispatch';
    var history = req.body.history || [];
    var lang = req.body.lang || 'English';

    var isFrederic = /i am fred|i.m fred|this is fred|hello.*fred|frederic here|i am frederic|frederic speaking/i.test(text.trim());

    var systemPrompt =
      'You are CAPTION, an operations assistant for AKI, a UAE logistics company. ' +
      'Built by Azhar (Mohammed Azharuddin, Customer Service and Operations at AKI). ' +
      'Azhar reports to Mr. Frederic Fleureau, GM Supply Chain and Operations Consumer at AKI. ' +
      'LANGUAGE: Respond ONLY in ' + lang + '. ' +
      (isFrederic ? 'FREDERIC MODE: Address as boss. Questions outside data: say Boss that is outside my scope. ' : '') +
      'CONVERSATION: Ongoing conversation. Remember everything. Answer follow-ups naturally. ' +
      'HOW TO READ THE DATA: ' +
      'SCREEN_NOW = exactly what the screen shows right now with all active filters. Use this ONLY when the user asks about "now" / "current view" / doesn\'t name a specific month. ' +
      'If the user names a SPECIFIC month (e.g. "June rejection details") and that month is NOT the one currently on screen, do NOT ask them to apply the filter — you already have every month\'s numbers in the MONTHLY section below. Read that month\'s data directly and answer immediately with the real numbers, in the SAME reply, every time. Never say "please apply the filter" or "I need to see that data" — you already have it. Set action=filter so the screen catches up to match your answer, but the spoken answer must stand on its own regardless of what the screen does. ' +
      'SCREEN_NOW tRej = TOTAL REJECTIONS on screen. ' +
      'SCREEN_NOW tDel = DELIVERED on screen. ' +
      'SCREEN_NOW contrib = CONTRIBUTION TO REJECTION RATE shown on screen. ' +
      'SCREEN_NOW val = VALUE AT RISK on screen. ' +
      'SCREEN_NOW foodRej = food type rejection count. nonFoodRej = non-food type rejection count. ' +
      'When food or nonfood type filter is active: DELIVERED on screen shows total YTD food/nonfood deliveries (large number like 57237 or 68265). ' +
      'For the actual June food delivered: add DCV.del[June] + DCF.del[June] from ByORG Monthly section. ' +
      'For actual June nonfood delivered: add DGC.del[June] + DGS.del[June] + DSN.del[June] from ByORG Monthly section. ' +
      'MONTHLY section = data for each month without day filter. ' +
      'Days section under each month = day-by-day breakdown. ' +
      'ByORG section = per ORG stats. ownRate = ORG rejection rate. contribBadge = small % shown on ORG card. ' +
      'RATE: overall rate = tRej/(tRej+tDel)*100. When type filter active use contrib from SCREEN_NOW. ' +
      'IMPORTANT: When day filter + type filter (food/nonfood) are both active: ' +
      'reason and customer counts are estimates scaled from all-type data. ' +
      'They may not add up exactly to tRej. Always state tRej as the exact total. ' +
      'Do NOT sum up reason counts and claim that is the total — use tRej from SCREEN_NOW. ' +
      'Keep answers 2 to 3 sentences. Lead with the actual number/answer in the first sentence — busy managers are listening, not reading, so do not warm up with preamble. If data missing say please upload the file. ' +
      'For dispatch/driver questions use AllDrivers section. ' +
      'Phone numbers: say plus then digits in groups. ' +
      'ORDER BOOKING vs DISPATCH — THESE ARE DIFFERENT THINGS, NEVER MIX THEM UP: ' +
      'ORDER_BOOKING section = orders that were BOOKED/PLACED (the "Today\'s Order Booking" dashboard). Use this whenever the user says "booking", "booked", "order booking", or "placed". ' +
      'DISPATCH section = orders that were physically DISPATCHED/sent out to customers already (a separate warehouse operation, often reflecting orders booked on a PREVIOUS day). Use this only when the user says "dispatch", "dispatched", or "delivery/route" questions. ' +
      'If the user asks "today\'s order booking" or "what did we book today", answer from ORDER_BOOKING, never from DISPATCH — booking and dispatch can be completely different numbers since today\'s dispatch is often yesterday\'s booking catching up. ' +
      'Set action=filter to filter dashboard. Set action=navigate to go to another dashboard. ' +
      'REJECTION FILTER: If the user asks about a specific month, ORG, food/non-food type, or external/internal source ' +
      'on the rejection dashboard, set action=filter and put the plain keywords in action_detail — ' +
      'e.g. action_detail="june non-food external" or action_detail="month=6 food". Always answer using the SAME period the user asked about, not the full year, unless they said "YTD" or "all months". ' +
      'SALON: Salon sales/rejections = the DGC org (label it "Salon" when the user says "salon"). ' +
      'CONTEXT SOURCES: the data below has two parts. The first part (SCREEN_NOW, MONTHLY, etc.) reflects exactly what is on the user\'s screen right now, including active filters — use it for "what am I looking at" style questions. ' +
      'The second part (sections ending in _DB, e.g. REJECTION_DB, DISPATCH_DB_LATEST, ORDER_BOOKING_DB, RETURNS_DB, BACKLOG_DB, AUTOMATION_DB, DELIVERY_DB, TEAM_DB) is pulled directly from the database every time and is ALWAYS complete and current, regardless of which dashboard tab the user currently has open or what they\'ve loaded this session. If a dashboard-specific section from the first part is missing or says "no file", fall back to the matching _DB section before ever saying data is unavailable — only say data is missing if the _DB section for that topic also says "no file uploaded". ' +
      'ALL DATA: ' + context.substring(0, 24000) +
      ' Reply ONLY in JSON: {"answer":"your answer","action":"none or filter or navigate","action_detail":"value","action_label":"label"}';

    var messages = [];
    var recent = history.slice(-16);
    for (var i = 0; i < recent.length; i++) {
      messages.push({ role: recent[i].role === 'assistant' ? 'assistant' : 'user', content: recent[i].content });
    }
    messages.push({ role: 'user', content: text });

    var msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 600,
      system: systemPrompt,
      messages: messages
    });

    var raw = (msg.content[0].text || '').trim();
    var parsed;
    try {
      var m2 = raw.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(m2 ? m2[0] : raw);
    } catch(e) {
      parsed = { answer: raw, action: 'none', action_label: '' };
    }
    res.json({ success: true, result: parsed });
  } catch(e) {
    console.error('/api/voice error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/chat', function(req, res) {
  try {
    var prompt=req.body.prompt, history=req.body.history||[];
    var messages=history.slice(-10).map(function(h){return{role:h.role==='assistant'?'assistant':'user',content:h.content};});
    if(!messages.length||messages[messages.length-1].content!==prompt) messages.push({role:'user',content:prompt});
    anthropic.messages.create({model:'claude-haiku-4-5-20251001',max_tokens:2000,system:'You are AZHAR-AI, a professional executive assistant for a UAE logistics company.',messages:messages})
      .then(function(msg){res.json({result:msg.content[0].text});})
      .catch(function(e){res.status(500).json({error:e.message});});
  } catch(e){if(!res.headersSent)res.status(500).json({error:e.message});}
});

app.post('/api/excel', upload.single('file'), function(req, res) {
  try {
    var question=req.body.question||'Analyse this data', dataText='';
    if(req.file){var ext2=path.extname(req.file.originalname||'').toLowerCase();dataText=(ext2==='.xlsx'||ext2==='.xls')?XLSX.utils.sheet_to_csv(XLSX.read(req.file.buffer,{type:'buffer'}).Sheets[XLSX.read(req.file.buffer,{type:'buffer'}).SheetNames[0]]):req.file.buffer.toString('utf8');}
    anthropic.messages.create({model:'claude-haiku-4-5-20251001',max_tokens:2000,messages:[{role:'user',content:question+(dataText?'\n\nData:\n'+dataText.substring(0,8000):'')}]})
      .then(function(msg){res.json({result:msg.content[0].text});})
      .catch(function(e){res.status(500).json({error:e.message});});
  } catch(e){if(!res.headersSent)res.status(500).json({error:e.message});}
});


// ─────────────────────────────────────────────
//  RETURNS DATA  (DB + file fallback)
// ─────────────────────────────────────────────
var returnsData = null;
var RETURNS_FILE = path.join(DATA_DIR, 'returns.json');

async function dbSaveReturns(uploadedBy, fileName, totalOrders, summary) {
  try {
    await pool.query('DELETE FROM returns_data');
    await pool.query(
      'INSERT INTO returns_data (uploaded_by, file_name, total_orders, summary) VALUES ($1, $2, $3, $4)',
      [uploadedBy, fileName, totalOrders, JSON.stringify(summary)]
    );
    return true;
  } catch(e) {
    console.error('DB save returns error:', e.message);
    return false;
  }
}

async function loadReturnsFromDB() {
  try {
    var res = await pool.query('SELECT * FROM returns_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      returnsData = {
        uploadedAt: res.rows[0].uploaded_at,
        uploadedBy: res.rows[0].uploaded_by,
        fileName: res.rows[0].file_name,
        totalOrders: res.rows[0].total_orders,
        summary: res.rows[0].summary
      };
      console.log('Loaded returns from DB');
      return true;
    }
  } catch(e) { console.error('DB load returns:', e.message); }
  var saved = loadJSON(RETURNS_FILE);
  if (saved) { returnsData = saved; console.log('Loaded returns from file'); }
  return false;
}
loadReturnsFromDB();

app.post('/api/returns/upload', async function(req, res) {
  try {
    var summary = (typeof req.body.summary === 'string') ? JSON.parse(req.body.summary || '{}') : (req.body.summary || {});
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var fileName = req.body.fileName || (req.file && req.file.originalname) || 'returns.csv';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    returnsData = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: uploadedBy,
      fileName: fileName,
      totalOrders: totalOrders,
      summary: summary
    };
    var dbOk = await dbSaveReturns(uploadedBy, fileName, totalOrders, summary);
    saveJSON(RETURNS_FILE, returnsData);
    var summaryKeys = Object.keys(summary);
    console.log('Returns saved:', totalOrders, 'orders, summary keys:', summaryKeys, dbOk ? '(DB+file)' : '(file only)');
    res.json({ success: true });
  } catch(e) {
    console.error('Returns upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.delete('/api/returns/clear', async function(req, res) {
  try {
    await pool.query('DELETE FROM returns_data');
    returnsData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var RETURNS_SUMMARY_VERSION = 'v3';
app.get('/api/returns/status', function(req, res) {
  if (!returnsData) return res.json({ hasData: false });
  if (!returnsData.summary || returnsData.summary.version !== RETURNS_SUMMARY_VERSION) {
    return res.json({ hasData: false, reason: 'version_mismatch' });
  }
  res.json({
    hasData: true,
    uploadedAt: returnsData.uploadedAt,
    uploadedBy: returnsData.uploadedBy,
    fileName: returnsData.fileName,
    totalOrders: returnsData.totalOrders,
    summary: returnsData.summary
  });
});

// ─── AUTH SYSTEM ──────────────────────────────────────────────────────────

// Audit log helper
async function auditLog(userId, username, action, details, ip) {
  try {
    await pool.query(
      'INSERT INTO audit_log (user_id, username, action, details, ip_address) VALUES ($1,$2,$3,$4,$5)',
      [userId||null, username||'system', action, details||'', ip||'']
    );
  } catch(e) { console.error('Audit log error:', e.message); }
}

// Auth middleware
async function requireAuth(req, res, next) {
  var token = req.headers['x-auth-token'] || req.headers['authorization'];
  if (token && token.startsWith('Bearer ')) token = token.slice(7);
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try {
    var sess = await pool.query(
      'SELECT s.*, u.id as uid, u.username, u.role, u.dashboards, u.full_name, u.active, u.must_change_password, u.horeca_salesperson_name FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=$1 AND s.expires_at>NOW()',
      [token]
    );
    if (!sess.rows[0]) return res.status(401).json({ error: 'Session expired' });
    if (!sess.rows[0].active) return res.status(403).json({ error: 'Account disabled' });
    req.user = sess.rows[0];
    next();
  } catch(e) { res.status(500).json({ error: e.message }); }
}

function requireRole(...roles) {
  return function(req, res, next) {
    if (!req.user) return res.status(401).json({ error: 'Not authenticated' });
    if (!roles.includes(req.user.role)) return res.status(403).json({ error: 'Access denied' });
    next();
  };
}

// ── HoReCa Order Module (new tables, shares this pool/auth, isolated from other dashboards) ──
require('./horeca_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);

// ── PPT Polish Module (upload a rough deck, get it improved via Claude, keeping original images) ──
require('./ppt_polish_module')(app, requireAuth, upload, anthropic, auditLog);

// ── Brand Presentations Module (CEO-level brand decks) ──
require('./brand_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);

// ── Vehicle Master Module (fleet registry, used as fallback for transport cost estimate) ──
var VEHICLE_MASTER_MAP = {};
require('./vehicle_master_module')(app, pool, requireAuth, requireRole, upload, auditLog, VEHICLE_MASTER_MAP);

// ── Aujan Pallet Collection & Recovery Tracking Module (Phase 1) ──
require('./pallet_module')(app, pool, requireAuth, requireRole, upload, auditLog);

// ── LOGIN ──
app.post('/api/auth/login', loginLimiter, async function(req, res) {
  try {
    var { username, password } = req.body;
    if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
    var result = await pool.query('SELECT * FROM users WHERE username=$1', [username.toLowerCase().trim()]);
    var user = result.rows[0];
    if (!user) return res.status(401).json({ error: 'Invalid username or password' });
    if (!user.active) return res.status(403).json({ error: 'Account is disabled. Contact admin.' });
    var match = await bcrypt.compare(password, user.password_hash);
    if (!match) return res.status(401).json({ error: 'Invalid username or password' });
    // Create session token
    var token = crypto.randomBytes(32).toString('hex');
    var expires = new Date(Date.now() + 8 * 60 * 60 * 1000); // 8 hours
    var ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || '';
    await pool.query('INSERT INTO sessions (token, user_id, expires_at, ip_address) VALUES ($1,$2,$3,$4)',
      [token, user.id, expires, ip]);
    // Update last login
    await pool.query('UPDATE users SET last_login=NOW() WHERE id=$1', [user.id]);
    // Audit
    await auditLog(user.id, user.username, 'LOGIN', 'Successful login', ip);
    res.json({
      success: true,
      token: token,
      user: {
        id: user.id,
        username: user.username,
        full_name: user.full_name,
        role: user.role,
        dashboards: user.dashboards,
        must_change_password: user.must_change_password
      }
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── LOGOUT ──
app.post('/api/auth/logout', requireAuth, async function(req, res) {
  try {
    var token = req.headers['x-auth-token'] || (req.headers['authorization']||'').replace('Bearer ','');
    await pool.query('DELETE FROM sessions WHERE token=$1', [token]);
    await auditLog(req.user.uid, req.user.username, 'LOGOUT', '', req.headers['x-forwarded-for']||'');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── GET ME ──
app.get('/api/auth/me', requireAuth, function(req, res) {
  res.json({
    id: req.user.uid,
    username: req.user.username,
    full_name: req.user.full_name,
    role: req.user.role,
    dashboards: req.user.dashboards,
    must_change_password: req.user.must_change_password
  });
});

// ── CHANGE PASSWORD ──
app.post('/api/auth/change-password', requireAuth, async function(req, res) {
  try {
    var { current_password, new_password } = req.body;
    if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
    var result = await pool.query('SELECT password_hash FROM users WHERE id=$1', [req.user.uid]);
    var match = await bcrypt.compare(current_password, result.rows[0].password_hash);
    if (!match) return res.status(401).json({ error: 'Current password incorrect' });
    var hash = await bcrypt.hash(new_password, 10);
    await pool.query('UPDATE users SET password_hash=$1, must_change_password=false WHERE id=$2', [hash, req.user.uid]);
    await auditLog(req.user.uid, req.user.username, 'CHANGE_PASSWORD', 'Password changed by user', '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── USER MANAGEMENT (superadmin only) ──
app.get('/api/users', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var result = await pool.query('SELECT id, username, full_name, role, dashboards, active, created_at, last_login, must_change_password FROM users ORDER BY created_at');
    res.json({ users: result.rows });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/users', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { username, password, full_name, role, dashboards } = req.body;
    if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
    var hash = await bcrypt.hash(password, 10);
    var dbs = dashboards || ['dispatch','rejection','summary','email','invoice','backlog','returns','sales','automation'];
    var result = await pool.query(
      'INSERT INTO users (username, password_hash, full_name, role, dashboards, created_by) VALUES ($1,$2,$3,$4,$5,$6) RETURNING id',
      [username.toLowerCase().trim(), hash, full_name||username, role||'viewer', JSON.stringify(dbs), req.user.username]
    );
    await auditLog(req.user.uid, req.user.username, 'CREATE_USER', 'Created user: '+username+' role: '+role, '');
    res.json({ success: true, id: result.rows[0].id });
  } catch(e) {
    if (e.message.includes('unique')) return res.status(400).json({ error: 'Username already exists' });
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/users/:id', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { full_name, role, dashboards, active } = req.body;
    var dbs = dashboards ? JSON.stringify(dashboards) : null;
    await pool.query(
      'UPDATE users SET full_name=COALESCE($1,full_name), role=COALESCE($2,role), dashboards=COALESCE($3::jsonb,dashboards), active=COALESCE($4,active) WHERE id=$5',
      [full_name||null, role||null, dbs, active!=null?active:null, req.params.id]
    );
    await auditLog(req.user.uid, req.user.username, 'UPDATE_USER', 'Updated user ID: '+req.params.id, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/users/:id', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    if (parseInt(req.params.id) === req.user.uid) return res.status(400).json({ error: 'Cannot delete your own account' });
    var u = await pool.query('SELECT username FROM users WHERE id=$1', [req.params.id]);
    await pool.query('DELETE FROM users WHERE id=$1', [req.params.id]);
    await auditLog(req.user.uid, req.user.username, 'DELETE_USER', 'Deleted user: '+(u.rows[0]||{}).username, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/users/:id/reset-password', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { new_password } = req.body;
    if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
    var hash = await bcrypt.hash(new_password, 10);
    await pool.query('UPDATE users SET password_hash=$1, must_change_password=true WHERE id=$2', [hash, req.params.id]);
    var u = await pool.query('SELECT username FROM users WHERE id=$1', [req.params.id]);
    await auditLog(req.user.uid, req.user.username, 'RESET_PASSWORD', 'Reset password for: '+(u.rows[0]||{}).username, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ── EMERGENCY ADMIN RESET (remove after first use) ──
app.get('/api/version', function(req, res) {
  res.json({ version: 'V4.1', date: '2026-06-05', status: 'running', auth: 'active' });
});

app.get('/api/setup/reset-admin', async function(req, res) {
  try {
    var hash = await bcrypt.hash('YAmaha100@', 10);
    var check = await pool.query("SELECT id FROM users WHERE username='azhar'");
    if (check.rows.length === 0) {
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role, active) VALUES ($1,$2,$3,$4,true)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      res.json({ success: true, message: 'Admin user CREATED. Login: azhar / YAmaha100@' });
    } else {
      // Force reset — update password AND ensure active=true
      await pool.query("UPDATE users SET password_hash=$1, active=true, must_change_password=false WHERE username='azhar'", [hash]);
      // Clear all old sessions for this user
      var uid = check.rows[0].id;
      await pool.query("DELETE FROM sessions WHERE user_id=$1", [uid]);
      res.json({ success: true, message: 'Password RESET. All sessions cleared. Login: azhar / YAmaha100@' });
    }
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── AUDIT LOG ──
app.get('/api/audit', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var result = await pool.query('SELECT * FROM audit_log ORDER BY created_at DESC LIMIT 500');
    res.json({ logs: result.rows });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── AUDIT UPLOAD ACTIONS ──
// Patch existing upload endpoints to log actions
// (handled via middleware injection in each upload route)

// STATIC - MUST BE LAST
app.get('/', function(req, res) {
  var p1=path.join(__dirname,'public','index.html'), p2=path.join(__dirname,'index.html'), p3=path.join(__dirname,'azhar-ai-v4.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  if(fs.existsSync(p3))return res.sendFile(p3);
  res.status(404).json({error:'index.html not found'});
});
app.get('/orders', function(req, res) {
  var p1=path.join(__dirname,'public','orders.html'), p2=path.join(__dirname,'orders.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'orders.html not found'});
});
app.get('/brands', function(req, res) {
  var p1=path.join(__dirname,'public','brand_frontend.html'), p2=path.join(__dirname,'brand_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'brand_frontend.html not found'});
});
app.get('/pallets', function(req, res) {
  var p1=path.join(__dirname,'public','pallet_frontend.html'), p2=path.join(__dirname,'pallet_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'pallet_frontend.html not found'});
});
app.use(express.static(path.join(__dirname,'public')));
app.use(express.static(__dirname));

app.use(function(err,req,res,next){
  console.error('Global error:',err.message);
  if(!res.headersSent)res.status(500).json({error:err.message||'Server error'});
});

// ─── DAILY SALES ──────────────────────────────────────────
var salesData = null;

async function loadSalesFromDB() {
  try {
    await pool.query('CREATE TABLE IF NOT EXISTS sales_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), uploaded_by TEXT, file_name TEXT, total_orders INT, summary JSONB)');
    var res = await pool.query('SELECT * FROM sales_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      salesData = { uploadedAt: res.rows[0].uploaded_at, fileName: res.rows[0].file_name, totalOrders: res.rows[0].total_orders, summary: res.rows[0].summary };
      console.log('Loaded sales from DB');
    }
  } catch(e) { console.error('DB load sales:', e.message); }
}
loadSalesFromDB();

app.delete('/api/sales/clear', async function(req, res) {
  try { await pool.query('DELETE FROM sales_data'); salesData = null; res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/sales/status', function(req, res) {
  if (!salesData) return res.json({ hasData: false });
  if (!salesData.summary || salesData.summary.version !== 'v1') return res.json({ hasData: false });
  res.json({ hasData: true, uploadedAt: salesData.uploadedAt, fileName: salesData.fileName, totalOrders: salesData.totalOrders, summary: salesData.summary });
});

app.post('/api/sales/upload', async function(req, res) {
  try {
    var summary = (typeof req.body.summary === 'string') ? JSON.parse(req.body.summary) : (req.body.summary || {});
    var fileName = req.body.fileName || 'sales.xlsx';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    salesData = { uploadedAt: new Date(), fileName: fileName, totalOrders: totalOrders, summary: summary };
    try {
      await pool.query('CREATE TABLE IF NOT EXISTS sales_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), uploaded_by TEXT, file_name TEXT, total_orders INT, summary JSONB)');
      await pool.query('DELETE FROM sales_data');
      await pool.query('INSERT INTO sales_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)', ['Admin', fileName, totalOrders, JSON.stringify(summary)]);
    } catch(dbErr) { console.error('Sales DB save:', dbErr.message); }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ─── GENERAL INFO ──────────────────────────────────────────────────────────
var genInfoData = null;
async function loadGenInfoFromDB() {
  try {
    await pool.query('CREATE TABLE IF NOT EXISTS geninfo_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), file_name TEXT, total_members INT, rows JSONB)');
    await pool.query(`CREATE TABLE IF NOT EXISTS automation_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_records INT,
      summary JSONB,
      rows JSONB
    )`);
    var r = await pool.query('SELECT * FROM geninfo_data ORDER BY uploaded_at DESC LIMIT 1');
    if (r.rows[0]) {
      genInfoData = { fileName: r.rows[0].file_name, totalMembers: r.rows[0].total_members, rows: r.rows[0].rows };
      console.log('GenInfo loaded:', genInfoData.fileName, genInfoData.totalMembers, 'members');
    }
  } catch(e) { console.error('GenInfo DB load:', e.message); }
}
loadGenInfoFromDB();

app.get('/api/geninfo/status', requireAuth, function(req, res) {
  if (!genInfoData) return res.json({ hasData: false });
  res.json({ hasData: true, fileName: genInfoData.fileName, totalMembers: genInfoData.totalMembers, rows: genInfoData.rows });
});

app.post('/api/geninfo/upload', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { rows, fileName, totalMembers } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: 'No rows provided' });
    genInfoData = { rows, fileName: fileName || 'team.xlsx', totalMembers: totalMembers || rows.length };
    try {
      await pool.query('DELETE FROM geninfo_data');
      await pool.query('INSERT INTO geninfo_data (file_name, total_members, rows) VALUES ($1,$2,$3)',
        [genInfoData.fileName, genInfoData.totalMembers, JSON.stringify(rows)]);
      console.log('GenInfo saved to DB:', rows.length, 'members');
    } catch(dbErr) { console.error('GenInfo DB save:', dbErr.message); }
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'GenInfo: ' + genInfoData.fileName, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/geninfo/clear', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM geninfo_data');
    genInfoData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ─── TWILIO VOIP ────────────────────────────────────────────────────────────
// STATUS: READY — Set these 4 env vars in Render to activate:
//   TWILIO_ACCOUNT_SID   → from twilio.com console
//   TWILIO_AUTH_TOKEN    → from twilio.com console
//   TWILIO_PHONE_NUMBER  → your Twilio number e.g. +12015551234
//   TWILIO_TWIML_APP_SID → create TwiML App in Twilio console, set Voice URL to:
//                          https://azr-operations.com/api/voip/twiml

var TWILIO_CONFIGURED = !!(
  process.env.TWILIO_ACCOUNT_SID &&
  process.env.TWILIO_AUTH_TOKEN &&
  process.env.TWILIO_PHONE_NUMBER &&
  process.env.TWILIO_TWIML_APP_SID
);

if (TWILIO_CONFIGURED) {
  console.log('✅ Twilio VoIP: CONFIGURED and ready');
} else {
  console.log('⚠ Twilio VoIP: Not configured (set env vars to activate)');
}

// Check VoIP status + issue browser token
app.get('/api/voip/status', requireAuth, async function(req, res) {
  if (!TWILIO_CONFIGURED) return res.json({ configured: false, reason: 'Missing env vars' });
  var twilio;
  try { twilio = require('twilio'); } catch(e) {
    return res.json({ configured: false, reason: 'twilio package missing' });
  }
  try {
    var accountSid  = process.env.TWILIO_ACCOUNT_SID;
    var apiKey      = process.env.TWILIO_API_KEY;
    var apiSecret   = process.env.TWILIO_API_SECRET;
    var twimlAppSid = process.env.TWILIO_TWIML_APP_SID;
    var identity    = (req.user.username || 'azhar_user').replace(/[^a-zA-Z0-9_]/g, '_');

    // AccessToken with API Key — required for Twilio JS SDK v1.x and v2.x
    var AccessToken = twilio.jwt.AccessToken;
    var VoiceGrant  = AccessToken.VoiceGrant;
    var grant = new VoiceGrant({
      outgoingApplicationSid: twimlAppSid,
      incomingAllow: false
    });
    var token = new AccessToken(accountSid, apiKey, apiSecret, {
      identity: identity,
      ttl: 3600
    });
    token.addGrant(grant);
    var jwt = token.toJwt();
    console.log('✅ VoIP token generated for:', identity);
    res.json({ configured: true, token: jwt });
  } catch(e) {
    console.error('❌ VoIP token error:', e.message);
    res.json({ configured: false, error: e.message });
  }
});

// TwiML — tells Twilio what to do when call connects (dial out to real number)
app.post('/api/voip/twiml', function(req, res) {
  var to = req.body.To || req.query.To;
  var callerId = process.env.TWILIO_PHONE_NUMBER || '';
  res.set('Content-Type', 'text/xml');
  if (to && to.startsWith('+')) {
    // Dial with two-way audio bridge
    res.send('<?xml version="1.0" encoding="UTF-8"?>' +
      '<Response>' +
        '<Dial callerId="' + callerId + '" timeout="30" record="do-not-record">' +
          '<Number>' + to + '</Number>' +
        '</Dial>' +
      '</Response>');
  } else if (to && to.startsWith('client:')) {
    // Browser client call
    res.send('<?xml version="1.0" encoding="UTF-8"?>' +
      '<Response>' +
        '<Dial callerId="' + callerId + '">' +
          '<Client>' + to.replace('client:','') + '</Client>' +
        '</Dial>' +
      '</Response>');
  } else {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Say>Call configuration error.</Say></Response>');
  }
});

// Initiate outbound call from server (alternative method)
app.post('/api/voip/call', requireAuth, async function(req, res) {
  if (!TWILIO_CONFIGURED) return res.status(503).json({ error: 'VoIP not configured' });
  try {
    var { to, from_number, from_name } = req.body;
    console.log('VoIP call request - to:', to, 'from:', from_number);
    if (!to) return res.status(400).json({ error: 'No destination number' });

    // ── SECURITY: Only allow calls to registered General Info numbers ──
    if (genInfoData && genInfoData.rows && genInfoData.rows.length) {
      var cleanTo = to.replace(/\s+/g, '').replace(/^00/, '+');
      var allowed = genInfoData.rows.some(function(row) {
        var contact = String(row['CONTACT'] || row['contact'] || '').replace(/\s+/g, '');
        if (!contact) return false;
        if (!contact.startsWith('+')) contact = '+' + contact;
        return contact === cleanTo ||
               contact.replace('+','') === cleanTo.replace('+','') ||
               cleanTo.endsWith(contact.slice(-9));
      });
      if (!allowed) {
        await auditLog(req.user.uid, req.user.username, 'VOIP_BLOCKED', 'Blocked: ' + to, '');
        return res.status(403).json({ error: 'Number not registered in General Info.' });
      }
    }

    var twilio = require('twilio');
    var client = twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);
    var conferenceName = 'bridge-' + Date.now();
    var baseUrl = 'https://azr-operations.com';

    if (from_number && from_number.startsWith('+')) {
      // Call caller first — when they answer, Twilio dials the destination
      var encodedTo = encodeURIComponent(to);
      var c1 = await client.calls.create({
        to: from_number,
        from: process.env.TWILIO_PHONE_NUMBER,
        url: baseUrl + '/api/voip/dial?to=' + encodedTo
      });
      console.log('VoIP BRIDGE: ' + from_number + ' -> ' + to);
      await auditLog(req.user.uid, req.user.username, 'VOIP_BRIDGE', from_number + ' <-> ' + to, '');
      res.json({ success: true, mode: 'bridge' });
    } else {
      // Outbound only
      var call = await client.calls.create({
        to: to,
        from: process.env.TWILIO_PHONE_NUMBER,
        url: baseUrl + '/api/voip/twiml'
      });
      console.log('VoIP OUTBOUND: ' + req.user.username + ' -> ' + to);
      await auditLog(req.user.uid, req.user.username, 'VOIP_CALL', 'Called: ' + to, '');
      res.json({ success: true, mode: 'outbound' });
    }
  } catch(e) {
    console.error('VoIP call error FULL:', JSON.stringify(e));
    console.error('VoIP call error msg:', e.message);
    console.error('TWILIO_ACCOUNT_SID set:', !!process.env.TWILIO_ACCOUNT_SID);
    console.error('TWILIO_AUTH_TOKEN set:', !!process.env.TWILIO_AUTH_TOKEN);
    console.error('TWILIO_PHONE_NUMBER set:', !!process.env.TWILIO_PHONE_NUMBER);
    res.status(500).json({ error: e.message });
  }
});

// ── CONFERENCE BRIDGE TwiML ──
app.get('/api/voip/dial', function(req, res) {
  var to = req.query.to || '';
  res.set('Content-Type', 'text/xml');
  if (to) {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial callerId="' + 
      (process.env.TWILIO_PHONE_NUMBER || '') + '" timeout="30">' + to + '</Dial></Response>');
  } else {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Say>Connection error.</Say></Response>');
  }
});

app.get('/api/voip/conference', function(req, res) {
  var room = req.query.room || 'azhar-default';
  res.set('Content-Type', 'text/xml');
  res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial><Conference startConferenceOnEnter="true" endConferenceOnExit="true" maxParticipants="2" record="do-not-record">' + room + '</Conference></Dial></Response>');
});

app.post('/api/voip/conference', function(req, res) {
  var room = req.query.room || req.body.room || 'azhar-default';
  res.set('Content-Type', 'text/xml');
  res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial><Conference startConferenceOnEnter="true" endConferenceOnExit="true" maxParticipants="2" record="do-not-record">' + room + '</Conference></Dial></Response>');
});

// ─── VOIP DEBUG (temporary) ─────────────────────────────────────────────────
app.get('/api/voip/debug', function(req, res) {
  res.json({
    configured: TWILIO_CONFIGURED,
    has_sid: !!process.env.TWILIO_ACCOUNT_SID,
    has_token: !!process.env.TWILIO_AUTH_TOKEN,
    has_phone: !!process.env.TWILIO_PHONE_NUMBER,
    has_twiml: !!process.env.TWILIO_TWIML_APP_SID,
    has_api_key: !!process.env.TWILIO_API_KEY,
    has_api_secret: !!process.env.TWILIO_API_SECRET,
    twilio_pkg: (function(){ try{ require('twilio'); return 'OK'; }catch(e){ return e.message; }})()
  });
});

// ─── SERVE TWILIO SDK ────────────────────────────────────────────────────────
app.get('/twilio-sdk.js', function(req, res) {
  try {
    // Try to serve from npm package
    var sdkPath = require.resolve('@twilio/voice-sdk/dist/twilio.js');
    res.setHeader('Content-Type', 'application/javascript');
    res.setHeader('Cache-Control', 'public, max-age=86400');
    require('fs').createReadStream(sdkPath).pipe(res);
  } catch(e) {
    try {
      var sdkPath2 = require.resolve('@twilio/voice-sdk/dist/twilio.min.js');
      res.setHeader('Content-Type', 'application/javascript');
      require('fs').createReadStream(sdkPath2).pipe(res);
    } catch(e2) {
      // Fallback: redirect to CDN (browser can access even if server can't)
      res.redirect('https://sdk.twilio.com/js/client/v1.14/twilio.js');
    }
  }
});

// ══════════════════════════════════════════════════════════
// AUTOMATION TRACKING API ROUTES — PER-MONTH MERGE
// ══════════════════════════════════════════════════════════
var automationData = null;

// New table structure: one row per month
// automation_months: month TEXT PK, auto INT, manual INT, total INT, org_data JSONB, updated_at, updated_by, file_name

async function initAutomationMonthsTable() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS automation_months (
      month TEXT PRIMARY KEY,
      auto_count INT DEFAULT 0,
      manual_count INT DEFAULT 0,
      total_count INT DEFAULT 0,
      org_data JSONB DEFAULT '{}'::jsonb,
      updated_at TIMESTAMPTZ DEFAULT NOW(),
      updated_by TEXT,
      file_name TEXT
    )`);
    console.log('automation_months table ready');
  } catch(e) { console.error('automation_months init:', e.message); }
}
initAutomationMonthsTable();

async function loadAutomationFromDB() {
  try {
    // Load from new per-month table
    var r = await pool.query('SELECT * FROM automation_months ORDER BY month');
    if (r.rows.length) {
      var monthData = {}, orgData = {}, totalAuto = 0, totalManual = 0;
      var latestFile = '', latestBy = '', latestAt = null;
      r.rows.forEach(function(row) {
        monthData[row.month] = { auto: row.auto_count, manual: row.manual_count, total: row.total_count };
        totalAuto   += row.auto_count;
        totalManual += row.manual_count;
        // Merge org data
        var od = row.org_data || {};
        Object.keys(od).forEach(function(org) {
          if (!orgData[org]) orgData[org] = { total:0, auto:0, manual:0, ots:[] };
          orgData[org].total  += od[org].total  || 0;
          orgData[org].auto   += od[org].auto   || 0;
          orgData[org].manual += od[org].manual || 0;
          (od[org].ots||[]).forEach(function(ot){
            if (orgData[org].ots.indexOf(ot) === -1) orgData[org].ots.push(ot);
          });
        });
        if (!latestAt || new Date(row.updated_at) > new Date(latestAt)) {
          latestAt = row.updated_at; latestFile = row.file_name; latestBy = row.updated_by;
        }
      });
      var total = totalAuto + totalManual;
      var rate  = total ? +(totalAuto/total*100).toFixed(2) : 0;
      var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var sortedMonths = months.filter(function(m){ return monthData[m]; });
      var latestMonth  = sortedMonths[sortedMonths.length-1] || '';
      var latestRate   = latestMonth ? +(monthData[latestMonth].auto/monthData[latestMonth].total*100).toFixed(2) : 0;
      automationData = {
        uploadedAt: latestAt, uploadedBy: latestBy, fileName: latestFile,
        totalRecords: total, rows: monthData, orgRows: orgData,
        summary: { total, auto: totalAuto, manual: totalManual, rate, orgRows: orgData },
        sortedMonths, latestMonth, latestRate
      };
      console.log('Automation loaded from DB:', total, 'total records across', sortedMonths.length, 'months:', sortedMonths.join(', '));
    }
  } catch(e) { console.error('Automation DB load:', e.message); }
}
loadAutomationFromDB();

app.get('/api/automation/status', requireAuth, function(req, res) {
  if (!automationData) return res.json({ hasData: false });
  res.json({
    hasData: true,
    uploadedAt:   automationData.uploadedAt,
    uploadedBy:   automationData.uploadedBy,
    fileName:     automationData.fileName,
    totalRecords: automationData.totalRecords,
    summary:      automationData.summary,
    rows:         automationData.rows,
    orgRows:      automationData.orgRows || {}
  });
});

app.post('/api/automation/upload', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { rows, orgRows, fileName, totalRecords, summary } = req.body;
    if (!rows) return res.status(400).json({ error: 'No data provided' });

    // rows = { Jan: {auto, manual, total}, Feb: {...}, ... }
    // orgRows = { 'Victory-Food': {total, auto, manual, ots:[...]}, ... }
    var monthsUpdated = [];

    for (var month in rows) {
      var md = rows[month];
      var od = {};
      // Build org_data for this month from orgRows — approximate split
      if (orgRows) {
        Object.keys(orgRows).forEach(function(org) {
          od[org] = orgRows[org]; // store full org totals per upload
        });
      }
      await pool.query(`
        INSERT INTO automation_months (month, auto_count, manual_count, total_count, org_data, updated_at, updated_by, file_name)
        VALUES ($1, $2, $3, $4, $5, NOW(), $6, $7)
        ON CONFLICT (month) DO UPDATE SET
          auto_count   = EXCLUDED.auto_count,
          manual_count = EXCLUDED.manual_count,
          total_count  = EXCLUDED.total_count,
          org_data     = EXCLUDED.org_data,
          updated_at   = NOW(),
          updated_by   = EXCLUDED.updated_by,
          file_name    = EXCLUDED.file_name
      `, [month, md.auto||0, md.manual||0, md.total||0, JSON.stringify(od), req.user.username, fileName||'automation.xlsx']);
      monthsUpdated.push(month);
    }

    console.log('Automation months upserted:', monthsUpdated.join(', '));
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Automation: ' + fileName + ' months: ' + monthsUpdated.join(','), '');

    // Reload full aggregated data from DB
    await loadAutomationFromDB();
    res.json({ success: true, monthsUpdated: monthsUpdated, totalRecords: automationData ? automationData.totalRecords : totalRecords });
  } catch(e) {
    console.error('Automation upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/automation/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM automation_months');
    automationData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Per-month delete endpoint (optional future use)
app.delete('/api/automation/month/:month', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM automation_months WHERE month=$1', [req.params.month]);
    await loadAutomationFromDB();
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var PORT=process.env.PORT||3000;
// ══════════════════════════════════════════════════════════
// DELIVERY SCHEDULE COMPLIANCE API ROUTES
// ══════════════════════════════════════════════════════════
var deliveryScheduleLookup = null;
var deliveryData = null;

async function initDeliveryTables() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS delivery_schedule (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      customer_count INT,
      lookup JSONB
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS delivery_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    // Load existing data
    var sr = await pool.query('SELECT * FROM delivery_schedule ORDER BY uploaded_at DESC LIMIT 1');
    if (sr.rows.length) {
      // Convert string keys to integers after JSONB parse
      var rawLookup = sr.rows[0].lookup || {};
      deliveryScheduleLookup = {};
      Object.keys(rawLookup).forEach(function(k) {
        deliveryScheduleLookup[parseInt(k)] = rawLookup[k];
      });
      console.log('Delivery schedule loaded:', Object.keys(deliveryScheduleLookup).length, 'customers');
    }
    var dr = await pool.query('SELECT * FROM delivery_data ORDER BY uploaded_at DESC LIMIT 1');
    if (dr.rows.length) {
      deliveryData = { summary: dr.rows[0].summary, uploadedBy: dr.rows[0].uploaded_by, fileName: dr.rows[0].file_name, totalOrders: dr.rows[0].total_orders };
      console.log('Delivery data loaded:', dr.rows[0].total_orders, 'orders');
    }
  } catch(e) { console.error('Delivery init:', e.message); }
}
initDeliveryTables();

app.get('/api/delivery/status', requireAuth, function(req, res) {
  res.json({
    hasSchedule: !!deliveryScheduleLookup,
    scheduleCustomers: deliveryScheduleLookup ? Object.keys(deliveryScheduleLookup).length : 0,
    scheduleLookup: deliveryScheduleLookup || {},
    hasData: !!deliveryData,
    summary: deliveryData ? deliveryData.summary : null,
    fileName: deliveryData ? deliveryData.fileName : null,
    uploadedBy: deliveryData ? deliveryData.uploadedBy : null,
    totalOrders: deliveryData ? deliveryData.totalOrders : 0
  });
});

app.post('/api/delivery/schedule', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { lookup, fileName, customerCount } = req.body;
    if (!lookup) return res.status(400).json({ error: 'No schedule data' });
    // Store with integer keys in memory
    deliveryScheduleLookup = {};
    Object.keys(lookup).forEach(function(k) {
      deliveryScheduleLookup[parseInt(k)] = lookup[k];
    });
    await pool.query('DELETE FROM delivery_schedule');
    await pool.query('INSERT INTO delivery_schedule (uploaded_by, file_name, customer_count, lookup) VALUES ($1,$2,$3,$4)',
      [req.user.username, fileName||'schedule.xlsx', customerCount||0, JSON.stringify(lookup)]);
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Schedule: ' + fileName, '');
    res.json({ success: true, customerCount: customerCount });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/delivery/data', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { summary, fileName, totalOrders } = req.body;
    if (!summary) return res.status(400).json({ error: 'No data' });
    deliveryData = { summary, fileName: fileName||'oracle.xlsx', uploadedBy: req.user.username, totalOrders: totalOrders||0 };
    await pool.query('DELETE FROM delivery_data');
    await pool.query('INSERT INTO delivery_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)',
      [req.user.username, deliveryData.fileName, deliveryData.totalOrders, JSON.stringify(summary)]);
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Data: ' + fileName, '');
    res.json({ success: true, totalOrders: deliveryData.totalOrders });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Server-side Oracle classification (avoids browser freeze) ────────
app.post('/api/delivery/classify', requireAuth, requireRole('superadmin','subadmin'), upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    if (!deliveryScheduleLookup || Object.keys(deliveryScheduleLookup).length === 0) {
      return res.status(400).json({ error: 'No schedule loaded on server. Please upload schedule first.' });
    }

    console.log('DS Classify: Reading', req.file.originalname, req.file.size, 'bytes');
    var wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true, cellNF: false, cellHTML: false, cellFormula: false });

    // Find correct sheet
    var sheetName = wb.SheetNames.find(function(s){ return s.trim() === 'Data'; })
      || wb.SheetNames.find(function(s){ return s.toUpperCase().includes('MASTER'); })
      || wb.SheetNames[0];

    console.log('DS Classify: Using sheet', sheetName, 'of', wb.SheetNames);
    var data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    console.log('DS Classify: Rows', data.length);

    if (!data.length) return res.status(400).json({ error: 'No data rows found in file. Sheet: ' + sheetName });

    var keys = Object.keys(data[0]);
    var colSiteId  = keys.find(function(k){ return /site.?id/i.test(k); }) || '';
    var colDate    = keys.find(function(k){ return /^rsd$/i.test(k.trim()) || /request.?date/i.test(k) || /invoice.?date/i.test(k); }) || '';
    var colTemp    = keys.find(function(k){ return /ambient.*frozen|frozen.*ambient/i.test(k) || k.trim() === 'Ambient / Frozen'; }) || '';
    var colChannel = keys.find(function(k){ return /channel/i.test(k); }) || '';
    var colMonth   = keys.find(function(k){ return k.trim().toUpperCase() === 'MONTH'; }) || '';
    var colCust    = keys.find(function(k){ return /customer.?name|customer_name/i.test(k); }) || '';
    var colOrg     = keys.find(function(k){ return k.trim().toUpperCase() === 'ORG'; }) || '';
    var colOrderType = keys.find(function(k){ return /order.?type/i.test(k); }) || '';

    console.log('DS Classify columns: siteId=' + colSiteId + ' date=' + colDate + ' temp=' + colTemp + ' channel=' + colChannel + ' month=' + colMonth);

    var DAYS = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
    var monthData = {}, channelData = {}, dayData = {}, noSchedCusts = {}, oosCusts = {};
    var tempData = { Ambient:{scheduled:0,oos:0,noSched:0,total:0}, Frozen:{scheduled:0,oos:0,noSched:0,total:0} };
    var scheduled = 0, oos = 0, noSched = 0;

    data.forEach(function(row) {
      var siteRaw = row[colSiteId];
      var site = null;
      try { site = siteRaw ? parseInt(parseFloat(String(siteRaw))) : null; } catch(e){}

      var day = null;
      var dateRaw = row[colDate];
      if (dateRaw) {
        try {
          var dt = null;
          if (dateRaw instanceof Date) { dt = dateRaw; }
          else {
            var f = parseFloat(String(dateRaw));
            if (!isNaN(f) && f > 1000) dt = new Date(Math.round((f - 25569) * 86400 * 1000));
            else dt = new Date(String(dateRaw));
          }
          if (dt && !isNaN(dt.getTime())) day = DAYS[dt.getDay()];
        } catch(e){}
      }

      var temp    = String(row[colTemp]||'').trim().toUpperCase();
      // Derive month from date if no MONTH column
      var month = '';
      if (colMonth && row[colMonth]) {
        month = String(row[colMonth]).trim();
      } else if (dateRaw) {
        try {
          var mdt = null;
          if (dateRaw instanceof Date) mdt = dateRaw;
          else { var mf = parseFloat(String(dateRaw)); if (!isNaN(mf) && mf > 1000) mdt = new Date(Math.round((mf-25569)*86400*1000)); else mdt = new Date(String(dateRaw)); }
          if (mdt && !isNaN(mdt.getTime())) {
            var mNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            month = mNames[mdt.getMonth()] + '-' + String(mdt.getFullYear()).slice(2);
          }
        } catch(e){}
      }
      var channel = String(row[colChannel]||'').trim();
      var org     = String(row[colOrg]||'').trim();
      var cust    = String(row[colCust]||'').trim();
      // Get channel and org from schedule if not in oracle
      if ((!channel || !org) && site && deliveryScheduleLookup[site]) {
        var sl2 = deliveryScheduleLookup[site];
        if (!channel && sl2.accountType) channel = sl2.accountType;
        if (!org && sl2.org) org = sl2.org;
      }

      var status;
      if (!site || !deliveryScheduleLookup[site]) {
        status = 'No Schedule';
        if (site) {
          if (!noSchedCusts[site]) noSchedCusts[site] = { name: cust, orders: 0, months: {}, channel: channel, org: org };
          noSchedCusts[site].orders++;
          if (month) noSchedCusts[site].months[month] = 1;
        }
      } else {
        var sl = deliveryScheduleLookup[site];
        var day3 = day ? day.substring(0,3) : null;
        if (day3 && sl.days && (day3 in sl.days)) {
          var schedTemp = sl.days[day3].replace(/\s/g,'').toUpperCase();
          if (temp === 'FROZEN') {
            status = schedTemp.indexOf('FROZEN') !== -1 ? 'Scheduled' : 'Out of Schedule';
          } else {
            status = schedTemp ? 'Scheduled' : 'Out of Schedule';
          }
        } else {
          status = 'Out of Schedule';
        }
        if (status === 'Out of Schedule') {
          if (!oosCusts[site]) oosCusts[site] = { name: cust, orders: 0, schedDays: sl.days ? Object.keys(sl.days).join(', ') : '', orderedDay: day3||'?', channel: channel, org: org };
          oosCusts[site].orders++;
        }
      }

      if (status === 'Scheduled') scheduled++;
      else if (status === 'Out of Schedule') oos++;
      else noSched++;
      // Track by temperature
      var tempKey = (temp === 'FROZEN') ? 'Frozen' : 'Ambient';
      if (!tempData[tempKey]) tempData[tempKey] = {scheduled:0,oos:0,noSched:0,total:0};
      tempData[tempKey][status==='Scheduled'?'scheduled':status==='Out of Schedule'?'oos':'noSched']++;
      tempData[tempKey].total++;

      if (month) {
        if (!monthData[month]) monthData[month] = { scheduled:0, oos:0, noSched:0, total:0 };
        monthData[month][status==='Scheduled'?'scheduled':status==='Out of Schedule'?'oos':'noSched']++;
        monthData[month].total++;
      }
      if (channel) {
        if (!channelData[channel]) channelData[channel] = { scheduled:0, oos:0, noSched:0, total:0 };
        channelData[channel][status==='Scheduled'?'scheduled':status==='Out of Schedule'?'oos':'noSched']++;
        channelData[channel].total++;
      }
      if (day) {
        var d3 = day.substring(0,3);
        if (!dayData[d3]) dayData[d3] = { scheduled:0, oos:0, noSched:0, total:0 };
        dayData[d3][status==='Scheduled'?'scheduled':status==='Out of Schedule'?'oos':'noSched']++;
        dayData[d3].total++;
      }
    });

    var total = scheduled + oos + noSched;
    var sp = total ? Math.round(scheduled/total*100) : 0;
    var op = total ? Math.round(oos/total*100) : 0;
    var np = total ? Math.round(noSched/total*100) : 0;

    var noSchedArr = Object.keys(noSchedCusts).map(function(s){
      var d = noSchedCusts[s];
      return { site:s, name:d.name, orders:d.orders, months:Object.keys(d.months).join(', '), channel:d.channel, org:d.org };
    }).sort(function(a,b){ return b.orders-a.orders; }).slice(0,50);

    var oosArr = Object.keys(oosCusts).map(function(s){
      var d = oosCusts[s];
      return { site:s, name:d.name, orders:d.orders, schedDays:d.schedDays, orderedDay:d.orderedDay, channel:d.channel, org:d.org };
    }).sort(function(a,b){ return b.orders-a.orders; }).slice(0,50);

    var summary = {
      total:total, scheduled:scheduled, oos:oos, noSched:noSched,
      schedPct:sp, oosPct:op, noSchedPct:np,
      monthData:monthData, channelData:channelData, dayData:dayData, tempData:tempData,
      noSchedCustomers:noSchedArr, oosCustomers:oosArr
    };

    // Save to DB
    await pool.query('DELETE FROM delivery_data');
    await pool.query('INSERT INTO delivery_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)',
      [req.user.username, req.file.originalname, total, JSON.stringify(summary)]);
    deliveryData = { summary, fileName: req.file.originalname, uploadedBy: req.user.username, totalOrders: total };

    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Oracle: ' + req.file.originalname + ' ' + total + ' orders', '');
    console.log('DS Classify complete:', total, 'orders — Scheduled:', sp + '%', 'OOS:', op + '%', 'NoSched:', np + '%');

    res.json({ success: true, summary: summary, totalOrders: total });
  } catch(e) {
    console.error('DS Classify error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/delivery/schedule/clear', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM delivery_schedule');
    deliveryScheduleLookup = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT,function(){console.log('AZHAR-AI server running on port '+PORT+(process.env.DATABASE_URL?' with PostgreSQL':' file-only mode'));});
