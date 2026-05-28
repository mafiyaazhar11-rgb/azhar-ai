const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── PERSIST PATHS ─────────────────────────────────────────────
const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DISPATCH_FILE  = path.join(DATA_DIR, 'dispatch.json');
const REJECTION_FILE = path.join(DATA_DIR, 'rejection.json');

function saveJSON(filepath, data) {
  try { fs.writeFileSync(filepath, JSON.stringify(data)); } catch (e) { console.error('Save error:', e.message); }
}
function loadJSON(filepath) {
  try { if (fs.existsSync(filepath)) return JSON.parse(fs.readFileSync(filepath, 'utf8')); } catch (e) {}
  return null;
}

app.get('/health', (req, res) => res.json({ status: 'ok', time: new Date().toISOString() }));

// ─────────────────────────────────────────────────────────────
// DISPATCH PARSER — FIXED TYPE COLUMN
// TYPE values in your file: Food / Non Food / 3PL / Van
// ─────────────────────────────────────────────────────────────
function normaliseType(raw) {
  if (!raw) return 'unknown';
  const t = raw.toString().trim().toUpperCase().replace(/\s+/g, ' ');
  if (t === 'FOOD')     return 'food';
  if (t === 'NON FOOD') return 'nonfood';
  if (t === '3PL' || t === '3 PL') return '3pl';
  if (t === 'VAN')      return 'van';
  return t.toLowerCase();
}

function parseDispatchData(input) {
  // input = Buffer (xlsx/xls) or string (csv/txt)
  let rows = [];

  if (Buffer.isBuffer(input)) {
    const wb = XLSX.read(input, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  } else {
    // CSV string — parse manually so we keep header names
    const lines = input.split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length < 2) return null;
    const headers = parseCSVLine(lines[0]);
    for (let i = 1; i < lines.length; i++) {
      const vals = parseCSVLine(lines[i]);
      if (!vals.length || !vals[0]) continue;
      const row = {};
      headers.forEach((h, idx) => { row[h.trim()] = (vals[idx] || '').trim(); });
      rows.push(row);
    }
  }

  if (!rows.length) return null;

  // Detect column names (case-insensitive)
  const sampleKeys = Object.keys(rows[0]).map(k => k.toUpperCase());
  const findCol = (...names) => {
    for (const n of names) {
      const found = Object.keys(rows[0]).find(k => k.toUpperCase().includes(n.toUpperCase()));
      if (found) return found;
    }
    return null;
  };

  const COL_ROUTE    = findCol('ROUTE');
  const COL_CITY     = findCol('CITY');
  const COL_CUSTOMER = findCol('CUSTOMER');
  const COL_AMOUNT   = findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE');
  const COL_DRIVER   = findCol('DRIVER CONTACT', 'DRIVER_CONTACT', 'DRIVER');
  const COL_LOCATION = findCol('LOCATION_ID', 'LOCATION');
  const COL_TYPE     = findCol('TYPE');
  const COL_ORG      = findCol('ORG');
  const COL_ORDER    = findCol('ORDER CODE', 'ORDER_CODE', 'ORDER');

  console.log('Detected columns:', { COL_TYPE, COL_ORG, COL_ROUTE, COL_CITY, COL_AMOUNT });

  // Counters
  let totalOrders = 0, totalValue = 0;
  let foodOrders = 0, foodValue = 0;
  let nonFoodOrders = 0, nonFoodValue = 0;
  let plOrders = 0, vanOrders = 0;

  const cities = {};
  const customers = {};
  const routes = {};
  const orgStats = {
    DCV: { o: 0, v: 0 }, DCF: { o: 0, v: 0 },
    DGC: { o: 0, v: 0 }, DGS: { o: 0, v: 0 },
    DSN: { o: 0, v: 0 }, HCP: { o: 0, v: 0 }
  };
  const driverSet = new Set();

  for (const row of rows) {
    const amt = parseFloat(row[COL_AMOUNT]) || 0;
    totalOrders++;
    totalValue += amt;

    // ── TYPE (most important fix) ──────────────────────────
    const rawType = COL_TYPE ? String(row[COL_TYPE] || '') : '';
    const type = normaliseType(rawType);

    if (type === 'food') {
      foodOrders++;
      foodValue += amt;
    } else if (type === 'nonfood') {
      nonFoodOrders++;
      nonFoodValue += amt;
    } else if (type === '3pl') {
      plOrders++;
    } else if (type === 'van') {
      vanOrders++;
    }

    // ── ORG ───────────────────────────────────────────────
    const rawOrg = COL_ORG ? String(row[COL_ORG] || '').trim().toUpperCase() : '';
    if (orgStats[rawOrg]) {
      orgStats[rawOrg].o++;
      orgStats[rawOrg].v += amt;
    } else if (rawOrg === '3 PL' || rawOrg === 'HCP') {
      orgStats.HCP.o++;
      orgStats.HCP.v += amt;
    }

    // ── CITY ──────────────────────────────────────────────
    if (COL_CITY && row[COL_CITY]) {
      const city = normaliseCity(String(row[COL_CITY] || ''));
      if (!cities[city]) cities[city] = { orders: 0, value: 0 };
      cities[city].orders++;
      cities[city].value += amt;
    }

    // ── CUSTOMER ──────────────────────────────────────────
    if (COL_CUSTOMER && row[COL_CUSTOMER]) {
      const cust = String(row[COL_CUSTOMER] || '').trim();
      if (!customers[cust]) customers[cust] = { orders: 0, value: 0 };
      customers[cust].orders++;
      customers[cust].value += amt;
    }

    // ── ROUTES & DRIVERS ──────────────────────────────────
    if (COL_ROUTE && row[COL_ROUTE]) {
      const route = String(row[COL_ROUTE] || '').trim();
      if (!routes[route]) routes[route] = { locations: new Set(), driver: '', value: 0 };
      const locId = COL_LOCATION ? String(row[COL_LOCATION] || '').trim() : '';
      if (locId) routes[route].locations.add(locId);
      routes[route].value += amt;

      // Extract driver name from contact field
      if (COL_DRIVER && row[COL_DRIVER] && !routes[route].driver) {
        routes[route].driver = extractDriverName(String(row[COL_DRIVER] || ''));
      }
    }

    // Track unique drivers
    if (COL_DRIVER && row[COL_DRIVER]) {
      const drvKey = extractDriverName(row[COL_DRIVER]) || String(row[COL_DRIVER] || '').trim();
      if (drvKey) driverSet.add(drvKey);
    }
  }

  // Debug log to verify counts
  console.log(`TYPE counts — food:${foodOrders} nonfood:${nonFoodOrders} 3pl:${plOrders} van:${vanOrders} total:${totalOrders}`);

  // ── AGGREGATE ─────────────────────────────────────────────
  const byCity = Object.entries(cities)
    .map(([city, v]) => ({ city, orders: v.orders, value: Math.round(v.value) }))
    .sort((a, b) => b.orders - a.orders);

  // Merge customer branches
  const baseCust = {};
  for (const [name, v] of Object.entries(customers)) {
    const base = stripBranch(name);
    if (!baseCust[base]) baseCust[base] = { orders: 0, value: 0 };
    baseCust[base].orders += v.orders;
    baseCust[base].value  += v.value;
  }
  const topCustomers = Object.entries(baseCust)
    .map(([name, v]) => ({ name, orders: v.orders, value: Math.round(v.value) }))
    .sort((a, b) => b.value - a.value).slice(0, 6);

  const topRoutes = Object.entries(routes)
    .map(([route, v]) => ({
      route,
      drops: v.locations.size,
      driver: v.driver,
      value: Math.round(v.value)
    }))
    .sort((a, b) => b.drops - a.drops).slice(0, 30);

  // Top drivers by drops
  const driverDrops = {};
  for (const [route, v] of Object.entries(routes)) {
    const drv = v.driver;
    if (!drv) continue;
    if (!driverDrops[drv]) driverDrops[drv] = 0;
    driverDrops[drv] += v.locations.size;
  }
  const topDrivers = Object.entries(driverDrops)
    .map(([name, orders]) => ({ name, orders }))
    .sort((a, b) => b.orders - a.orders).slice(0, 5);

  const totalDrops = Object.values(routes).reduce((s, r) => s + r.locations.size, 0);

  return {
    total_orders:    totalOrders,
    total_value:     Math.round(totalValue),
    total_routes:    Object.keys(routes).length,
    total_drivers:   driverSet.size || Object.keys(routes).length,
    total_drops:     totalDrops,
    food_orders:     foodOrders,
    food_value:      Math.round(foodValue),
    non_food_orders: nonFoodOrders,
    non_food_value:  Math.round(nonFoodValue),
    pl_orders:       plOrders,
    van_orders:      vanOrders,
    type_breakdown: {
      DCV: { orders: orgStats.DCV.o, value: Math.round(orgStats.DCV.v) },
      DCF: { orders: orgStats.DCF.o, value: Math.round(orgStats.DCF.v) },
      DGC: { orders: orgStats.DGC.o, value: Math.round(orgStats.DGC.v) },
      DGS: { orders: orgStats.DGS.o, value: Math.round(orgStats.DGS.v) },
      DSN: { orders: orgStats.DSN.o, value: Math.round(orgStats.DSN.v) },
      HCP: { orders: orgStats.HCP.o, value: Math.round(orgStats.HCP.v) }
    },
    by_city:       byCity,
    top_customers: topCustomers,
    top_drivers:   topDrivers,
    top_routes:    topRoutes
  };
}

// ── HELPERS ───────────────────────────────────────────────────
function parseCSVLine(line) {
  const result = [];
  let cell = '', inQ = false;
  for (const ch of line) {
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ',' && !inQ) { result.push(cell); cell = ''; }
    else { cell += ch; }
  }
  result.push(cell);
  return result;
}

function normaliseCity(raw) {
  const c = (raw || '').trim().toLowerCase();
  if (c.includes('abu dhabi'))  return 'Abu Dhabi';
  if (c.includes('dubai'))      return 'Dubai';
  if (c.includes('sharjah'))    return 'Sharjah';
  if (c.includes('ajman'))      return 'Ajman';
  if (c.includes('fujairah'))   return 'Fujairah';
  if (c.includes('al ain') || c.includes('al-ain') || c === 'alain') return 'Al Ain';
  if (c.includes('ras al') || c === 'rak') return 'Ras Al Khaimah';
  if (c.includes('umm'))        return 'Umm Al Quwain';
  const rawStr = String(raw || '').trim();
  return rawStr.charAt(0).toUpperCase() + rawStr.slice(1).toLowerCase();
}

function extractDriverName(contact) {
  if (!contact) return '';
  const str = String(contact).trim();
  const match = str.match(/^([A-Za-z][A-Za-z\s]{1,30?}?)(?:\s*[-+\d])/);
  if (match) return match[1].trim();
  const parts = str.split(/[-+\d]/);
  return (parts[0] || '').trim();
}

function stripBranch(name) {
  let base = String(name || '');
  const stripAfter = [',Branch', ', Branch', ',Br.', ', Br.', ' -Branch', ',CPD', ' CPD', '- Branch', '-Branch'];
  for (const kw of stripAfter) {
    const i = base.toLowerCase().indexOf(kw.toLowerCase());
    if (i > 3) { base = base.substring(0, i).trim(); break; }
  }
  return base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
}

// ─────────────────────────────────────────────────────────────
// DISPATCH MEMORY  (persisted to disk)
// ─────────────────────────────────────────────────────────────
let dispatchHistory = {};
let currentDispatch = null;

const savedDispatch = loadJSON(DISPATCH_FILE);
if (savedDispatch) {
  dispatchHistory = savedDispatch.history || {};
  const keys = Object.keys(dispatchHistory).sort().reverse();
  if (keys.length > 0) currentDispatch = dispatchHistory[keys[0]];
  console.log('Loaded dispatch history:', keys.length, 'dates');
}

// ── DISPATCH ROUTES ───────────────────────────────────────────
app.post('/api/dispatch/upload', upload.single('file'), async (req, res) => {
  try {
    let summary = null;
    let csvText = '';

    if (req.file) {
      const ext = path.extname(req.file.originalname || '').toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        summary  = parseDispatchData(req.file.buffer);
        // Also keep CSV for AI queries
        const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
        csvText  = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
      } else {
        csvText  = req.file.buffer.toString('utf8');
        summary  = parseDispatchData(csvText);
      }
    } else if (req.body.csvText) {
      csvText = req.body.csvText;
      summary = parseDispatchData(csvText);
    }

    if (!summary) return res.status(400).json({ error: 'No data parsed from file. Check file format.' });

    const dateKey    = req.body.dateKey    || new Date().toISOString().split('T')[0];
    const uploadedBy = req.body.uploadedBy || 'Admin';

    const entry = {
      uploadedAt: new Date().toISOString(),
      uploadedBy,
      csvText: csvText.substring(0, 200000), // keep for AI, cap size
      summary,
      date: dateKey
    };

    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;

    // Keep only last 30 dates
    const keys = Object.keys(dispatchHistory).sort();
    while (keys.length > 30) { delete dispatchHistory[keys.shift()]; }

    saveJSON(DISPATCH_FILE, { history: dispatchHistory });
    console.log(`Dispatch saved: ${dateKey} — ${summary.total_orders} orders, AED ${summary.total_value}`);

    res.json({ success: true, summary, uploadedAt: entry.uploadedAt, date: dateKey });
  } catch (e) {
    console.error('Dispatch upload error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/dispatch/status', (req, res) => {
  const availableDates = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData: false, availableDates });
  res.json({
    hasData: true,
    uploadedAt:  currentDispatch.uploadedAt,
    uploadedBy:  currentDispatch.uploadedBy,
    summary:     currentDispatch.summary,
    date:        currentDispatch.date,
    availableDates
  });
});

app.get('/api/dispatch/date/:dateKey', (req, res) => {
  const entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData: false });
  currentDispatch = entry;
  res.json({
    hasData:    true,
    uploadedAt: entry.uploadedAt,
    uploadedBy: entry.uploadedBy,
    summary:    entry.summary,
    date:       entry.date
  });
});

app.post('/api/dispatch/ask', async (req, res) => {
  try {
    if (!currentDispatch) {
      return res.json({ result: 'No dispatch data loaded. Please upload today\'s report.' });
    }
    const s = currentDispatch.summary;
    const context = `
Dispatch Date: ${currentDispatch.date}
Total Orders: ${s.total_orders}
Total Value: AED ${s.total_value?.toLocaleString()}
Food Orders: ${s.food_orders} orders | AED ${s.food_value?.toLocaleString()}
Non-Food Orders: ${s.non_food_orders} orders | AED ${s.non_food_value?.toLocaleString()}
3PL Orders: ${s.pl_orders}
Van Orders: ${s.van_orders || 0}
Total Routes: ${s.total_routes}
Total Drivers: ${s.total_drivers}
Total Drops: ${s.total_drops}

ORG Breakdown:
${Object.entries(s.type_breakdown || {}).map(([k, v]) => `  ${k}: ${v.orders} orders, AED ${v.value?.toLocaleString()}`).join('\n')}

Top Cities:
${(s.by_city || []).slice(0, 8).map(c => `  ${c.city}: ${c.orders} orders`).join('\n')}

Top Customers:
${(s.top_customers || []).map((c, i) => `  ${i+1}. ${c.name}: ${c.orders} orders, AED ${c.value?.toLocaleString()}`).join('\n')}

Top Drivers:
${(s.top_drivers || []).map((d, i) => `  ${i+1}. ${d.name}: ${d.orders} drops`).join('\n')}

Top Routes:
${(s.top_routes || []).slice(0, 10).map(r => `  ${r.route} | ${r.driver || '—'} | ${r.drops} drops | AED ${r.value?.toLocaleString()}`).join('\n')}

CSV Sample (first 8000 chars):
${currentDispatch.csvText?.substring(0, 8000) || ''}`;

    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      messages: [{
        role: 'user',
        content: `You are AZHAR-AI Dispatch Intelligence for UAE logistics.\n\n${context}\n\nQuestion: ${req.body.question}\n\nAnswer with exact numbers. Use AED for currency. Be concise and precise.`
      }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    console.error('Dispatch ask error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─────────────────────────────────────────────────────────────
// REJECTION STORE  (persisted to disk)
// ─────────────────────────────────────────────────────────────
let rejectionData = null;

const savedRejection = loadJSON(REJECTION_FILE);
if (savedRejection) {
  rejectionData = savedRejection;
  console.log('Loaded rejection data — uploadedAt:', rejectionData.uploadedAt);
}

app.post('/api/rejection/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls') {
      return res.status(400).json({ error: 'Please upload .xlsx or .xls file' });
    }

    const wb   = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!rows.length) return res.status(400).json({ error: 'No rows found in file' });

    // Detect column names
    const sampleKeys = Object.keys(rows[0]);
    const findRejCol = (...names) => sampleKeys.find(k => names.some(n => k.toUpperCase().includes(n.toUpperCase()))) || null;

    const ROOT_COL   = findRejCol('ROOT CAUSE', 'ROOT_CAUSE', 'FINA') || 'fina- Root Cause';
    const ORG_COL    = findRejCol('ORGANIZATION', 'ORG') || 'Organization';
    const DATE_COL   = findRejCol('D DATE', 'DATE', 'DELIVERY DATE') || 'D DATE';
    const STATUS_COL = findRejCol('STATUS') || 'Status';
    const CUST_COL   = findRejCol('CUSTOMER NAME', 'CUSTOMER') || 'CUSTOMER NAME';
    const AREA_COL   = findRejCol('AREA', 'CITY') || 'Area';
    const VALUE_COL  = findRejCol('VALUE', 'AMOUNT') || 'Value';
    const TYPE_COL   = findRejCol('TYPE') || null;

    console.log('Rejection cols:', { ROOT_COL, ORG_COL, DATE_COL, STATUS_COL });

    const isRej = (row) => {
      const s = String(row[STATUS_COL] || '').trim().toUpperCase();
      return s === 'R/D' || s === 'HOLD' || s === 'RD' || s === 'REJECTED' || s === 'R';
    };
    const isDel = (row) => {
      const s = String(row[STATUS_COL] || '').trim().toUpperCase();
      return s.includes('DELIVER') || s === 'D' || s === 'DELIVERED';
    };

    const parseDate = (v) => {
      if (!v) return null;
      if (v instanceof Date) return v;
      if (typeof v === 'number') {
        // Excel serial date
        const d = XLSX.SSF.parse_date_code(v);
        if (d) return new Date(d.y, d.m - 1, d.d);
      }
      const d = new Date(v);
      return isNaN(d) ? null : d;
    };

    const orgMap   = {};
    const monthMap = {};
    let totalRej = 0, totalDel = 0, totalVal = 0;

    for (const row of rows) {
      const rej = isRej(row);
      const del = isDel(row);
      if (!rej && !del) continue;

      const d    = parseDate(row[DATE_COL]);
      const mo   = d ? d.getMonth() + 1 : null;
      const day  = d ? d.getDate() : null;
      const org  = String(row[ORG_COL]  || '').trim().toUpperCase();
      const root = String(row[ROOT_COL] || '').trim();
      const cust = String(row[CUST_COL] || '').trim();
      const area = String(row[AREA_COL] || '').trim();
      const val  = parseFloat(row[VALUE_COL]) || 0;
      const type = TYPE_COL ? normaliseType(row[TYPE_COL] || '') : 'unknown';

      if (del) totalDel++;
      if (rej) { totalRej++; totalVal += val; }

      // ORG level
      if (org) {
        if (!orgMap[org]) orgMap[org] = {
          tDel: 0, tRej: 0, val: 0,
          del: [0,0,0,0,0,0,0,0,0,0,0,0],
          rej: [0,0,0,0,0,0,0,0,0,0,0,0],
          reasons: {}, custs: {}, areas: {}
        };
        if (del) { orgMap[org].tDel++; if (mo) orgMap[org].del[mo-1]++; }
        if (rej) {
          orgMap[org].tRej++;
          orgMap[org].val += val;
          if (mo) orgMap[org].rej[mo-1]++;
          if (root) orgMap[org].reasons[root] = (orgMap[org].reasons[root] || 0) + 1;
          if (cust) orgMap[org].custs[cust]   = (orgMap[org].custs[cust]   || 0) + 1;
          if (area) orgMap[org].areas[area]   = (orgMap[org].areas[area]   || 0) + 1;
        }
      }

      // Month/Day level
      if (mo) {
        if (!monthMap[mo]) monthMap[mo] = {
          days: new Set(), tDel: 0, tRej: 0, val: 0, reasons: {}, data: {}
        };
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++;
          monthMap[mo].val += val;
          if (root) monthMap[mo].reasons[root] = (monthMap[mo].reasons[root] || 0) + 1;
          if (day) monthMap[mo].days.add(day);
        }
        if (day) {
          if (!monthMap[mo].data[day]) monthMap[mo].data[day] = {
            tDel: 0, tRej: 0, val: 0, reasons: {}, custs: {}, areas: {}
          };
          if (del) monthMap[mo].data[day].tDel++;
          if (rej) {
            monthMap[mo].data[day].tRej++;
            monthMap[mo].data[day].val += val;
            if (root) monthMap[mo].data[day].reasons[root] = (monthMap[mo].data[day].reasons[root] || 0) + 1;
            if (cust) monthMap[mo].data[day].custs[cust]   = (monthMap[mo].data[day].custs[cust]   || 0) + 1;
            if (area) monthMap[mo].data[day].areas[area]   = (monthMap[mo].data[day].areas[area]   || 0) + 1;
          }
        }
      }
    }

    const fmtVal = (v) => v >= 1000000 ? 'AED ' + (v / 1000000).toFixed(2) + 'M' : 'AED ' + Math.round(v / 1000) + 'K';
    const top5   = (obj) => Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([l, n]) => ({ l, n }));
    const top5c  = (obj) => Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([n, c]) => ({ n, c, v: '' }));
    const top6a  = (obj) => Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([a, n]) => ({ a, n }));

    // Build ALL combined
    const allReasons = {}, allCusts = {}, allAreas = {};
    const allDel = [0,0,0,0,0,0,0,0,0,0,0,0];
    const allRej = [0,0,0,0,0,0,0,0,0,0,0,0];
    for (const v of Object.values(orgMap)) {
      for (const [k, n] of Object.entries(v.reasons)) allReasons[k] = (allReasons[k] || 0) + n;
      for (const [k, n] of Object.entries(v.custs))   allCusts[k]   = (allCusts[k]   || 0) + n;
      for (const [k, n] of Object.entries(v.areas))   allAreas[k]   = (allAreas[k]   || 0) + n;
      v.del.forEach((d, i) => allDel[i] += d);
      v.rej.forEach((r, i) => allRej[i] += r);
    }

    // Serialise months (Set → Array)
    const monthsOut = {};
    for (const [mo, md] of Object.entries(monthMap)) {
      const dataOut = {};
      for (const [day, dd] of Object.entries(md.data)) {
        dataOut[day] = {
          tDel: dd.tDel, tRej: dd.tRej, val: fmtVal(dd.val),
          reasons: top5(dd.reasons), custs: top5c(dd.custs), areas: top6a(dd.areas)
        };
      }
      monthsOut[mo] = {
        days:    Array.from(md.days).sort((a, b) => a - b),
        tDel:    md.tDel,
        tRej:    md.tRej,
        val:     fmtVal(md.val),
        reasons: top5(md.reasons),
        data:    dataOut
      };
    }

    // Serialise orgs
    const orgsOut = {
      all: {
        tDel:    totalDel,
        tRej:    totalRej,
        val:     fmtVal(totalVal),
        del:     allDel,
        rej:     allRej,
        reasons: top5(allReasons),
        custs:   top5c(allCusts),
        areas:   top6a(allAreas)
      }
    };
    for (const [org, v] of Object.entries(orgMap)) {
      orgsOut[org] = {
        tDel:    v.tDel,
        tRej:    v.tRej,
        val:     fmtVal(v.val),
        del:     v.del,
        rej:     v.rej,
        reasons: top5(v.reasons),
        custs:   top5c(v.custs),
        areas:   top6a(v.areas)
      };
    }

    rejectionData = {
      uploadedAt:  new Date().toISOString(),
      uploadedBy:  req.body.uploadedBy || 'Admin',
      fileName:    req.file.originalname,
      totalOrders: totalRej + totalDel,
      orgs:        orgsOut,
      months:      monthsOut
    };

    saveJSON(REJECTION_FILE, rejectionData);
    console.log(`Rejection saved: ${totalRej} rejections, ${totalDel} delivered`);

    res.json({ success: true, summary: { totalRej, totalDel, fileName: req.file.originalname } });
  } catch (e) {
    console.error('Rejection upload error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

// CRITICAL: This must return JSON always — never HTML
app.get('/api/rejection/status', (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  if (!rejectionData) return res.json({ hasData: false });
  res.json({
    hasData:     true,
    uploadedAt:  rejectionData.uploadedAt,
    uploadedBy:  rejectionData.uploadedBy,
    fileName:    rejectionData.fileName,
    totalOrders: rejectionData.totalOrders,
    orgs:        rejectionData.orgs,
    months:      rejectionData.months
  });
});

// ─── GENERAL CHAT ─────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    let messages = [];
    if (history && history.length > 0) {
      messages = history.slice(-10).map(h => ({
        role: h.role === 'assistant' ? 'assistant' : 'user',
        content: h.content
      }));
      if (!messages.length || messages[messages.length - 1].content !== prompt) {
        messages.push({ role: 'user', content: prompt });
      }
    } else {
      messages = [{ role: 'user', content: prompt }];
    }
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 2000,
      system: 'You are AZHAR-AI, a professional executive assistant for a UAE logistics company. Be concise and helpful.',
      messages
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    console.error('Chat error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── EXCEL ANALYSIS ───────────────────────────────────────────
app.post('/api/excel', upload.single('file'), async (req, res) => {
  try {
    let question = req.body.question || 'Analyse this data';
    let dataText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname || '').toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
        dataText = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
      } else {
        dataText = req.file.buffer.toString('utf8');
      }
    }
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 2000,
      messages: [{
        role: 'user',
        content: `${question}\n\n${dataText ? 'Data:\n' + dataText.substring(0, 8000) : ''}`
      }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    console.error('Excel error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── POWERPOINT ───────────────────────────────────────────────
app.post('/api/powerpoint', upload.single('file'), async (req, res) => {
  try {
    const { topic, slides, tone } = req.body;
    let extra = '';
    if (req.file) extra = '\n\nSource content:\n' + req.file.buffer.toString('utf8').substring(0, 3000);
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 3000,
      messages: [{
        role: 'user',
        content: `Create a ${slides || 5}-slide ${tone || 'professional'} presentation about: ${topic}${extra}\n\nFormat each slide as:\nSlide 1: [Title]\n- Bullet point\n\nSpeaker Notes: [notes]`
      }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    console.error('Powerpoint error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ─── STATIC FILES — must be LAST, after all API routes ──────
app.get('/', (req, res) => {
  const p1 = path.join(__dirname, 'public', 'index.html');
  const p2 = path.join(__dirname, 'index.html');
  const p3 = path.join(__dirname, 'azhar-ai-v4.html');
  if (fs.existsSync(p1)) return res.sendFile(p1);
  if (fs.existsSync(p2)) return res.sendFile(p2);
  if (fs.existsSync(p3)) return res.sendFile(p3);
  res.status(404).json({ error: 'index.html not found' });
});
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// ─── START ────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✦ AZHAR-AI server running on port ${PORT}`));
