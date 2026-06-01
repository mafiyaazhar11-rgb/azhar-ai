const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 150 * 1024 * 1024 } });

app.use(cors());
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
    // AUTH TABLES
    await pool.query(`CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      full_name TEXT,
      role TEXT NOT NULL DEFAULT 'viewer',
      dashboards JSONB DEFAULT '["dispatch","rejection","summary","email","invoice","backlog","returns","sales"]'::jsonb,
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
    // Create default super admin if not exists
    var adminCheck = await pool.query("SELECT id FROM users WHERE username = 'azhar'");
    if (adminCheck.rows.length === 0) {
      var hash = await bcrypt.hash('YAmaha100@', 10);
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role) VALUES ($1,$2,$3,$4)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      console.log('Default super admin created: azhar / azhar2026');
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
    // Keep only latest rejection upload (upsert by overwriting)
    await pool.query(`DELETE FROM rejection_data`);
    await pool.query(`
      INSERT INTO rejection_data (uploaded_by, file_name, total_orders, orgs, months)
      VALUES ($1, $2, $3, $4, $5)
    `, [uploadedBy, fileName, totalOrders, JSON.stringify(orgs), JSON.stringify(months)]);
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

function normaliseType(raw) {
  var t = toStr(raw).toUpperCase().replace(/\s+/g, ' ').trim();
  if (t === 'FOOD' || t.startsWith('FOOD')) return 'food';
  if (t.includes('NON-FOOD') || t.includes('NON FOOD')) return 'nonfood';
  if (t === 'GSEB' || t === 'SHARKNINJA') return 'nonfood';
  if (t === '3PL' || t === '3 PL') return '3pl';
  if (t === 'VAN') return 'van';
  return t.toLowerCase();
}

function normaliseCity(raw) {
  var c = toStr(raw).toLowerCase();
  if (c.includes('abu dhabi')) return 'Abu Dhabi';
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

function extractDriverName(contact) {
  var s = toStr(contact);
  if (!s) return '';
  if (/^[A-Za-z][A-Za-z\s]{2,}$/.test(s)) return s.trim();
  var m = s.match(/^([A-Za-z][A-Za-z\s]{2,29})(?:\s*[-+\d])/);
  if (m) return m[1].trim();
  var parts = s.split(/[-+\d]/);
  var name = (parts[0] || '').trim();
  return name.length > 2 ? name : s.trim();
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

  var C = {
    route:    findCol('ROUTE'),
    city:     findCol('CITY', 'AREA'),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount:   findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    driver:   findCol('DRIVER CONTACT DETAILS', 'DRIVERS NAME', 'DRIVER NAME', 'DRIVER CONTACT', 'DRIVER_CONTACT', 'DRIVER_ID'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    keep:     findCol('KEEP TOGETHER', 'KEEP_TOGETHER', 'KEEPTOGETHER', 'KEEP'),
    type:     findCol('TYPE'),
    org:      findCol('ORG') || findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU')
  };
  console.log('Dispatch cols:', JSON.stringify(C));

  var totalOrders=0, totalValue=0, foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  var cities={}, customers={}, routes={}, driverSet={};
  var orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, DPS:{o:0,v:0}, DPB:{o:0,v:0}, HCP:{o:0,v:0} };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    totalOrders++;
    var amt = parseFloat(row[C.amount]) || 0;
    totalValue += amt;
    var type = normaliseType(C.type ? row[C.type] : '');
    if (type === 'food')   { foodOrders++;    foodValue    += amt; }
    else if (type === 'nonfood') { nonFoodOrders++; nonFoodValue += amt; }
    else if (type === '3pl')     { plOrders++; }
    else if (type === 'van')     { vanOrders++; }
    var org = C.org ? toStr(row[C.org]).toUpperCase() : '';
    if (orgStats[org]) { orgStats[org].o++; orgStats[org].v += amt; }
    else if (org === '3 PL' || org === 'HCP' || org === '3PL') { orgStats.HCP.o++; orgStats.HCP.v += amt; }
    if (C.city && row[C.city]) {
      var city = normaliseCity(row[C.city]);
      if (!cities[city]) cities[city] = { orders:0, value:0 };
      cities[city].orders++; cities[city].value += amt;
    }
    if (C.customer && row[C.customer]) {
      var cust = toStr(row[C.customer]);
      if (!customers[cust]) customers[cust] = { orders:0, value:0 };
      customers[cust].orders++; customers[cust].value += amt;
    }
    if (C.route && row[C.route]) {
      var route = toStr(row[C.route]);
      if (!routes[route]) routes[route] = { locs:{}, drivers:{}, orders:0, value:0 };
      var loc = C.location ? toStr(row[C.location]) : '';
      if (loc) routes[route].locs[loc] = 1;
      routes[route].value += amt;
      routes[route].orders++;
      if (C.driver && row[C.driver]) {
        var drvName = extractDriverName(row[C.driver]);
        if (drvName) routes[route].drivers[drvName] = 1;
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
    return { route:route, orders:routes[route].orders, drops:Object.keys(routes[route].locs).length, driverCount:Object.keys(routes[route].drivers).length, value:Math.round(routes[route].value) };
  }).sort(function(a,b) { return b.drops-a.drops; }).slice(0,30);

  // Count actual orders per driver (not route drops)
  var driverOrders = {};
  rows.forEach(function(row) {
    var drv = C.driver && row[C.driver] ? extractDriverName(row[C.driver]) : '';
    if (!drv) return;
    var amt = C.amount ? parseFloat(row[C.amount]) || 0 : 0;
    var locId = C.location ? toStr(row[C.location]) : '';
    if (!driverOrders[drv]) driverOrders[drv] = {orders:0, drops:{}, value:0};
    driverOrders[drv].orders++;
    driverOrders[drv].value += amt;
    if (locId) driverOrders[drv].drops[locId] = 1;
  });
  var topDrivers = Object.keys(driverOrders).map(function(name) {
    return { name:name, orders:driverOrders[name].orders, drops:Object.keys(driverOrders[name].drops).length, value:Math.round(driverOrders[name].value) };
  }).sort(function(a,b) { return b.orders-a.orders; }).slice(0,5);

  return {
    total_orders: totalOrders, total_value: Math.round(totalValue),
    total_routes: Object.keys(routes).length,
    total_drivers: Object.keys(driverOrders).length || Object.keys(driverSet).length,
    total_drops: Object.keys(routes).reduce(function(s,r){ return s + Object.keys(routes[r].locs).length; }, 0),
    food_orders: foodOrders, food_value: Math.round(foodValue),
    non_food_orders: nonFoodOrders, non_food_value: Math.round(nonFoodValue),
    pl_orders: plOrders, van_orders: vanOrders,
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
    top_drivers: topDrivers, top_routes: topRoutes
  };
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
      rejectionData = {
        uploadedAt: row.uploaded_at, uploadedBy: row.uploaded_by,
        fileName: row.file_name, totalOrders: row.total_orders,
        orgs: row.orgs, months: row.months
      };
      console.log('Loaded rejection from DB uploadedAt:', rejectionData.uploadedAt);
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
      var csvText = req.file.buffer.toString('utf8');
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
      root:    findC('FINA- ROOT', 'ROOT CAUSE', 'ROOT_CAUSE', 'REASON-1'),
      cust:    findC('CUSTOMER NAME', 'CUSTOMER'),
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
      if (v instanceof Date) return v;
      if (typeof v === 'number') {
        try { var unix=Math.round((v-25569)*86400*1000); var dd=new Date(unix); if(!isNaN(dd.getTime()))return dd; } catch(e2){}
      }
      var d=new Date(v); return isNaN(d.getTime())?null:d;
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
      var root=toStr(row[RC.root]);
      var cust=toStr(row[RC.cust]);
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
        if (!orgMap[org]) orgMap[org]={tDel:0,tRej:0,val:0,food_rej:0,food_del:0,nonfood_rej:0,nonfood_del:0,ext_rej:0,ext_del:0,int_rej:0,int_del:0,food_val:0,nonfood_val:0,del:new Array(12).fill(0),rej:new Array(12).fill(0),reasons:{},custs:{},areas:{}};
        if (del) {
          orgMap[org].tDel++; if(mo)orgMap[org].del[mo-1]++;
          if(isFood)orgMap[org].food_del++; else if(isNF)orgMap[org].nonfood_del++;
          if(srcStr==='EXTERNAL')orgMap[org].ext_del++; else if(srcStr==='INTERNAL')orgMap[org].int_del++;
        }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val+=val; if(mo)orgMap[org].rej[mo-1]++;
          if(isFood){orgMap[org].food_rej++;orgMap[org].food_val+=val;}
          else if(isNF){orgMap[org].nonfood_rej++;orgMap[org].nonfood_val+=val;}
          if(srcStr==='EXTERNAL')orgMap[org].ext_rej++; else if(srcStr==='INTERNAL')orgMap[org].int_rej++;
          if(root)orgMap[org].reasons[root]=(orgMap[org].reasons[root]||0)+1;
          if(cust)orgMap[org].custs[cust]=(orgMap[org].custs[cust]||0)+1;
          if(area)orgMap[org].areas[area]=(orgMap[org].areas[area]||0)+1;
        }
      }
      if (mo) {
        if (!monthMap[mo]) monthMap[mo]={days:{},tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{},data:{}};
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++; monthMap[mo].val+=val;
          if(root)monthMap[mo].reasons[root]=(monthMap[mo].reasons[root]||0)+1;
          if(cust)monthMap[mo].custs[cust]=(monthMap[mo].custs[cust]||0)+1;
          if(area)monthMap[mo].areas[area]=(monthMap[mo].areas[area]||0)+1;
          if(day)monthMap[mo].days[day]=1;
        }
        if (day) {
          if(!monthMap[mo].data[day])monthMap[mo].data[day]={tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{}};
          if(del)monthMap[mo].data[day].tDel++;
          if(rej){monthMap[mo].data[day].tRej++;monthMap[mo].data[day].val+=val;if(root)monthMap[mo].data[day].reasons[root]=(monthMap[mo].data[day].reasons[root]||0)+1;if(cust)monthMap[mo].data[day].custs[cust]=(monthMap[mo].data[day].custs[cust]||0)+1;if(area)monthMap[mo].data[day].areas[area]=(monthMap[mo].data[day].areas[area]||0)+1;}
        }
      }
    }

    function fmtVal(v){return v>=1000000?'AED '+(v/1000000).toFixed(2)+'M':'AED '+Math.round(v/1000)+'K';}
    function top10(obj){return Object.keys(obj).map(function(l){return{l:l,n:obj[l]};}).sort(function(a,b){return b.n-a.n;}).slice(0,10);}
    function top8c(obj){return Object.keys(obj).map(function(n){return{n:n,c:obj[n],v:''};}).sort(function(a,b){return b.c-a.c;}).slice(0,8);}
    function top6a(obj){return Object.keys(obj).map(function(a){return{a:a,n:obj[a]};}).sort(function(a,b){return b.n-a.n;}).slice(0,6);}

    var allR={},allC={},allA={},allDel=new Array(12).fill(0),allRej=new Array(12).fill(0);
    var allFoodRej=0,allNFRej=0,allExtRej=0,allIntRej=0,allFoodDel=0,allNFDel=0,allFoodVal=0,allNFVal=0;
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      Object.keys(v.reasons).forEach(function(k){allR[k]=(allR[k]||0)+v.reasons[k];});
      Object.keys(v.custs).forEach(function(k){allC[k]=(allC[k]||0)+v.custs[k];});
      Object.keys(v.areas).forEach(function(k){allA[k]=(allA[k]||0)+v.areas[k];});
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
        dataOut[day]={tDel:dd.tDel,tRej:dd.tRej,val:fmtVal(dd.val),reasons:top10(dd.reasons),custs:top8c(dd.custs),areas:top6a(dd.areas)};
      });
      monthsOut[mo]={days:Object.keys(md.days).map(Number).sort(function(a,b){return a-b;}),tDel:md.tDel,tRej:md.tRej,val:fmtVal(md.val),reasons:top10(md.reasons),custs:top8c(md.custs||{}),areas:top6a(md.areas||{}),data:dataOut};
    });

    var orgsOut={all:{tDel:totalDel,tRej:totalRej,val:fmtVal(totalVal),food_rej:allFoodRej,food_del:allFoodDel,nonfood_rej:allNFRej,nonfood_del:allNFDel,ext_rej:allExtRej,int_rej:allIntRej,food_val:fmtVal(allFoodVal),nonfood_val:fmtVal(allNFVal),del:allDel,rej:allRej,reasons:top10(allR),custs:top8c(allC),areas:top6a(allA)}};
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      orgsOut[org]={tDel:v.tDel,tRej:v.tRej,val:fmtVal(v.val),food_rej:v.food_rej||0,food_del:v.food_del||0,nonfood_rej:v.nonfood_rej||0,nonfood_del:v.nonfood_del||0,ext_rej:v.ext_rej||0,int_rej:v.int_rej||0,food_val:fmtVal(v.food_val||0),nonfood_val:fmtVal(v.nonfood_val||0),del:v.del,rej:v.rej,reasons:top10(v.reasons),custs:top8c(v.custs),areas:top6a(v.areas)};
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
  res.json({ hasData:true, uploadedAt:rejectionData.uploadedAt, uploadedBy:rejectionData.uploadedBy, fileName:rejectionData.fileName, totalOrders:rejectionData.totalOrders, orgs:rejectionData.orgs, months:rejectionData.months });
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

app.post('/api/voice', requireAuth, async function(req, res) {
  try {
    var text = req.body.text || '';
    var context = req.body.context || '';
    var tab = req.body.tab || 'dispatch';

    var isGreeting = /^(hi|hello|hey|good morning|good evening|jarvis)/i.test(text.trim());
    var isFrederic = /i am fred|i.m fred|this is fred|hello.*fred|frederic here|i am frederic|frederic speaking/i.test(text.trim());
    var fredericMode = isFrederic || /boss|frederic/i.test(text.trim());

    var prompt =
      'You are JARVIS, a sharp and intelligent operations assistant for a UAE logistics company. ' +
      'You were built by Azhar — Mohammed Azharuddin from the Customer Service and Operations team at AKI. ' +
      'If anyone asks who built you: say I was built by Azhar, Mohammed Azharuddin from Customer Service and Operations at AKI. ' +
      'Azhar reports to Mr. Frederic Fleureau, General Manager Supply Chain and Operations Consumer at AKI. ' +
      '\n\nSPECIAL FREDERIC MODE RULES (apply ONLY when user says they are Frederic or mentions Frederic):' +
      '\n- If user says "I am Frederic" or identifies as Frederic: respond with "Yes boss! Welcome sir. Azhar speaks very highly of you. How can I assist you today, boss?"' +
      '\n- Always address Frederic as "boss" in every reply.' +
      '\n- If Frederic asks how JARVIS is: say "Fully operational boss. Always ready to serve. Azhar built me with your vision in mind."' +
      '\n- If Frederic asks about himself: say he is the General Manager Supply Chain and Operations Consumer at AKI, known for encouraging team initiative and setting high operational standards.' +
      '\n- If Frederic asks anything NOT related to the dashboard data: say "Boss, that is beyond my current scope. I will pass this to my boss Azhar and he will come back to you with what you need."' +
      '\n- If Frederic asks about Azhar: say Azhar is your dedicated operations analyst who built this platform to serve your vision boss. He is always working to improve it for you.' +
      '\n- Give Frederic warm professional compliments naturally — he is a great leader who inspires the team.' +
      '\nEND FREDERIC MODE RULES\n\n' +
      'Speak confidently like a male professional. Be direct — no fluff. Use exact numbers from data. ' +
      '\n\nCURRENT DASHBOARD: ' + tab +
      '\nDATA AVAILABLE:\n' + context.substring(0, 4000) +
      '\n\nUSER COMMAND: "' + text + '"' +
      '\n\nINSTRUCTIONS:' +
      '\n- CRITICAL: You are on the ' + tab + ' dashboard. Use ONLY data from that dashboard.' +
      '\n- ALWAYS use exact numbers from DATA AVAILABLE. Never say zero or not available if numbers exist.' +
      '\n- If greeting: say Hello, I am JARVIS your operations assistant. How can I help you today.' +
      '\n- Total sales/orders/value: use total_orders and total_value fields.' +
      '\n- Food sales: use food_orders and food_value fields.' +
      '\n- Non food sales: use non_food_orders and non_food_value fields.' +
      '\n- DCV/DGC/DCF/DGS/DSN/DPS sales: use type_breakdown.DCV.orders and type_breakdown.DCV.value etc.' +
      '\n- WH Backlog: look for category counts - advance orders, backlog orders, credit hold, frozen orders.' +
      '\n- Top customers: list name and AED value from top_customers array.' +
      '\n- Top drivers: list name, drops, and AED value from top_drivers array.' +
      '\n- If asking to filter by route/ORG/warehouse/city/BU: set action=filter.' +
      '\n- If asking to go to another dashboard: set action=navigate.' +
      '\n- If no data uploaded yet: say please upload the file first.' +
      '\n- Keep answer under 3 sentences. Use exact numbers. Speak like a confident male professional.' +
      '\n- GENERAL INFO: When asked for contact or phone number, read the PHONE field as a phone number, not as a numeric value. Always say it as a phone number.' +
      '\n- GENERAL INFO: To find who handles an outlet, check Outlets field and return name and PHONE.' +
      '\n\nReply ONLY with valid JSON, no markdown, no extra text:' +
      '\n{"answer":"precise answer with exact numbers","action":"none or filter or navigate","action_detail":"value","action_label":"action description"}';

    var msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 500,
      messages: [{ role: 'user', content: prompt }]
    });

    var raw = (msg.content[0].text || '').trim();
    var parsed;
    try {
      var match = raw.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(match ? match[0] : raw);
    } catch(e) {
      parsed = { answer: raw, action: 'none', action_label: '' };
    }
    res.json({ success: true, result: parsed });
  } catch(e) {
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
      'SELECT s.*, u.id as uid, u.username, u.role, u.dashboards, u.full_name, u.active, u.must_change_password FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=$1 AND s.expires_at>NOW()',
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

// ── LOGIN ──
app.post('/api/auth/login', async function(req, res) {
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
    var dbs = dashboards || ['dispatch','rejection','summary','email','invoice','backlog','returns','sales'];
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
app.get('/api/setup/reset-admin', async function(req, res) {
  try {
    var hash = await bcrypt.hash('YAmaha100@', 10);
    // Check if user exists
    var check = await pool.query("SELECT id FROM users WHERE username='azhar'");
    if (check.rows.length === 0) {
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role) VALUES ($1,$2,$3,$4)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      res.json({ success: true, message: 'Admin user created' });
    } else {
      await pool.query("UPDATE users SET password_hash=$1 WHERE username='azhar'", [hash]);
      res.json({ success: true, message: 'Admin password reset to YAmaha100@' });
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

var PORT=process.env.PORT||3000;
app.listen(PORT,function(){console.log('AZHAR-AI server running on port '+PORT+(process.env.DATABASE_URL?' with PostgreSQL':' file-only mode'));});
