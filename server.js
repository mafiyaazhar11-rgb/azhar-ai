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
app.use(express.static('public'));
app.use(express.static('.'));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── PERSIST PATHS ─────────────────────────────────────────────
const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DISPATCH_FILE   = path.join(DATA_DIR, 'dispatch.json');
const REJECTION_FILE  = path.join(DATA_DIR, 'rejection.json');

function saveJSON(filepath, data) {
  try { fs.writeFileSync(filepath, JSON.stringify(data)); } catch(e) { console.error('Save error:', e.message); }
}
function loadJSON(filepath) {
  try { if (fs.existsSync(filepath)) return JSON.parse(fs.readFileSync(filepath, 'utf8')); } catch(e) {}
  return null;
}

// Serve index.html
app.get('/', (req, res) => {
  const p1 = path.join(__dirname, 'public', 'index.html');
  const p2 = path.join(__dirname, 'index.html');
  if (fs.existsSync(p1)) res.sendFile(p1);
  else if (fs.existsSync(p2)) res.sendFile(p2);
  else res.send('AZHAR-AI Running');
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ─────────────────────────────────────────────────────────────
// DISPATCH PARSER
// ─────────────────────────────────────────────────────────────
function parseDispatchCSV(csvText, dateKey) {
  const lines = csvText.split('\n').filter(l => l.trim());
  if (lines.length < 2) return { total_orders: 0, date: dateKey };

  const headers = lines[0].split(',').map(h => h.trim().replace(/['"]/g,'').toUpperCase());
  const idx = (names) => {
    for (const n of names) { const i = headers.findIndex(h => h.includes(n)); if (i >= 0) return i; }
    return -1;
  };

  const routeIdx    = idx(['ROUTE']);
  const cityIdx     = idx(['CITY']);
  const custIdx     = idx(['CUSTOMER']);
  const amtIdx      = idx(['TOTAL_AMOUNT','AMOUNT']);
  const driverIdx   = idx(['DRIVER CONTACT','DRIVER_CONTACT']);
  const driverIdIdx = idx(['DRIVER_ID']);
  const orderIdx    = idx(['ORDER CODE','ORDER_CODE']);
  const locationIdx = idx(['LOCATION_ID','LOCATION']);
  const typeColIdx  = idx(['TYPE']);
  const orgColIdx   = idx(['ORG']);

  let totalOrders = 0, totalValue = 0;
  const cities = {}, customers = {}, routes = {}, driverSet = new Set();
  const orgStats = { DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, HCP:{o:0,v:0} };
  let foodOrders=0, foodValue=0, noonFoodOrders=0, noonFoodValue=0, plOrders=0, vanOrders=0;

  for (let i = 1; i < lines.length; i++) {
    const row = [];
    let cell = '', inQ = false;
    for (const ch of lines[i]) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { row.push(cell.trim()); cell = ''; }
      else { cell += ch; }
    }
    row.push(cell.trim());
    if (!row[0] || row[0] === '') continue;
    totalOrders++;

    const amt = amtIdx >= 0 ? parseFloat(row[amtIdx]) || 0 : 0;
    totalValue += amt;

    const rawType = typeColIdx >= 0 ? (row[typeColIdx] || '').trim() : '';
    const typeUpper = rawType.toUpperCase();
    if (typeUpper === 'FOOD') { foodOrders++; foodValue += amt; }
    else if (typeUpper === 'NOON FOOD') { noonFoodOrders++; noonFoodValue += amt; }
    else if (typeUpper === '3 PL' || typeUpper === '3PL') { plOrders++; }
    else if (typeUpper === 'VAN') { vanOrders++; }

    const rawOrg = orgColIdx >= 0 ? (row[orgColIdx] || '').trim().toUpperCase() : '';
    if (rawOrg === 'DCV')      { orgStats.DCV.o++; orgStats.DCV.v += amt; }
    else if (rawOrg === 'DCF') { orgStats.DCF.o++; orgStats.DCF.v += amt; }
    else if (rawOrg === 'DGC') { orgStats.DGC.o++; orgStats.DGC.v += amt; }
    else if (rawOrg === 'DGS') { orgStats.DGS.o++; orgStats.DGS.v += amt; }
    else if (rawOrg === 'DSN') { orgStats.DSN.o++; orgStats.DSN.v += amt; }
    else if (rawOrg === '3 PL' || rawOrg === 'HCP') { orgStats.HCP.o++; orgStats.HCP.v += amt; }

    if (cityIdx >= 0 && row[cityIdx]) {
      const rawCity = row[cityIdx].trim();
      let city = rawCity.charAt(0).toUpperCase() + rawCity.slice(1).toLowerCase();
      if (city.toLowerCase().includes('abu dhabi'))              city = 'Abu Dhabi';
      else if (city.toLowerCase().includes('ras al') || city.toLowerCase() === 'rak') city = 'Ras Al Khaimah';
      else if (city.toLowerCase() === 'al ain' || city.toLowerCase() === 'al-ain') city = 'Al Ain';
      else if (city.toLowerCase() === 'dubai')   city = 'Dubai';
      else if (city.toLowerCase() === 'sharjah') city = 'Sharjah';
      else if (city.toLowerCase() === 'ajman')   city = 'Ajman';
      else if (city.toLowerCase() === 'fujairah') city = 'Fujairah';
      if (!cities[city]) cities[city] = { orders: 0, value: 0 };
      cities[city].orders++;
      cities[city].value += amt;
    }

    if (custIdx >= 0 && row[custIdx]) {
      const cust = row[custIdx].trim();
      if (!customers[cust]) customers[cust] = { orders: 0, value: 0 };
      customers[cust].orders++;
      customers[cust].value += amt;
    }

    if (routeIdx >= 0 && row[routeIdx]) {
      const route = row[routeIdx].trim();
      const locId = locationIdx >= 0 ? row[locationIdx]?.trim() : '';
      let driverName = '';
      if (driverIdx >= 0 && row[driverIdx]) {
        const contact = row[driverIdx].trim();
        const match = contact.match(/^([A-Za-z][A-Za-z\s]+?)(?:[\-\+\s]+[0-9]|$)/);
        driverName = match ? match[1].trim() : contact.split(/[\-\+0-9]/)[0].trim();
      }
      if (!routes[route]) routes[route] = { locations: new Set(), driver: driverName, value: 0, orderLines: 0 };
      if (locId) routes[route].locations.add(locId);
      routes[route].orderLines++;
      routes[route].value += amt;
      if (driverName && !routes[route].driver) routes[route].driver = driverName;
    }

    if (driverIdIdx >= 0 && row[driverIdIdx]) driverSet.add(row[driverIdIdx].trim());
    else if (routeIdx >= 0 && row[routeIdx])  driverSet.add(row[routeIdx].trim());
  }

  const byCity = Object.entries(cities)
    .map(([city, v]) => ({ city, orders: v.orders, value: Math.round(v.value) }))
    .sort((a,b) => b.orders - a.orders);

  const baseCust = {};
  Object.entries(customers).forEach(([name, v]) => {
    let base = name;
    const stripAfter = [',Branch',', Branch',',Br.',', Br.',' -Branch',',CPD',' CPD','- Branch','-Branch'];
    for (const kw of stripAfter) {
      const i = base.toLowerCase().indexOf(kw.toLowerCase());
      if (i > 3) { base = base.substring(0, i).trim(); break; }
    }
    base = base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
    if (!baseCust[base]) baseCust[base] = { orders: 0, value: 0, locations: 0 };
    baseCust[base].orders   += v.orders;
    baseCust[base].value    += v.value;
    baseCust[base].locations++;
  });

  const topCustomers = Object.entries(baseCust)
    .map(([name, v]) => ({ name, orders: v.orders, value: Math.round(v.value), locations: v.locations }))
    .sort((a,b) => b.value - a.value).slice(0, 6);

  const topRoutes = Object.entries(routes)
    .map(([route, v]) => ({ route, drops: v.locations.size, order_lines: v.orderLines, driver: v.driver, value: Math.round(v.value) }))
    .sort((a,b) => b.drops - a.drops).slice(0, 30);

  // Serialize Sets for JSON storage
  const topRoutesSerial = topRoutes.map(r => ({ ...r }));

  const topDrivers = Object.entries(routes)
    .filter(([,v]) => v.driver)
    .map(([route, v]) => ({ name: v.driver, orders: v.locations.size, route }))
    .sort((a,b) => b.orders - a.orders).slice(0, 5);

  const luluEntries = Object.entries(customers).filter(([n]) => n.toUpperCase().includes('LULU'));
  const luluOrders  = luluEntries.reduce((s,[,v]) => s + v.orders, 0);
  const luluValue   = luluEntries.reduce((s,[,v]) => s + v.value,  0);

  return {
    total_orders:    totalOrders,
    total_value:     Math.round(totalValue),
    total_routes:    Object.keys(routes).length,
    total_drivers:   driverSet.size || Object.keys(routes).length,
    lulu_orders:     luluOrders,
    lulu_value:      Math.round(luluValue),
    food_orders:     foodOrders,
    food_value:      Math.round(foodValue),
    non_food_orders: noonFoodOrders,
    non_food_value:  Math.round(noonFoodValue),
    pl_orders:       plOrders,
    van_orders:      vanOrders,
    total_drops:     Object.values(routes).reduce((s, r) => s + (r.locations ? r.locations.size : 0), 0),
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
    top_routes:    topRoutesSerial,
    date:          dateKey
  };
}

// ─────────────────────────────────────────────────────────────
// DISPATCH MEMORY STORE  (persisted to disk)
// ─────────────────────────────────────────────────────────────
let dispatchHistory = {};
let currentDispatch = null;

// Load persisted dispatch on startup
const savedDispatch = loadJSON(DISPATCH_FILE);
if (savedDispatch) {
  dispatchHistory = savedDispatch.history || {};
  const keys = Object.keys(dispatchHistory).sort().reverse();
  if (keys.length > 0) currentDispatch = dispatchHistory[keys[0]];
  console.log('Loaded dispatch history:', keys.length, 'dates');
}

app.post('/api/dispatch/upload', upload.single('file'), async (req, res) => {
  try {
    let csvText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname || '').toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        csvText = XLSX.utils.sheet_to_csv(ws);
      } else {
        csvText = req.file.buffer.toString('utf8');
      }
    } else if (req.body.csvText) {
      csvText = req.body.csvText;
    }
    if (!csvText) return res.status(400).json({ error: 'No file received' });

    const dateKey    = req.body.dateKey    || new Date().toISOString().split('T')[0];
    const uploadedBy = req.body.uploadedBy || 'Admin';
    const summary    = parseDispatchCSV(csvText, dateKey);

    const entry = { uploadedAt: new Date().toISOString(), uploadedBy, csvText, summary, date: dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;

    const keys = Object.keys(dispatchHistory).sort();
    if (keys.length > 30) delete dispatchHistory[keys[0]];

    // Persist to disk
    saveJSON(DISPATCH_FILE, { history: dispatchHistory });

    res.json({ success: true, summary, uploadedAt: entry.uploadedAt, date: dateKey });
  } catch (e) {
    console.error('Dispatch upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/dispatch/status', (req, res) => {
  const availableDates = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData: false, availableDates });
  res.json({ hasData: true, uploadedAt: currentDispatch.uploadedAt, uploadedBy: currentDispatch.uploadedBy, summary: currentDispatch.summary, date: currentDispatch.date, availableDates });
});

app.get('/api/dispatch/date/:dateKey', (req, res) => {
  const entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData: false });
  currentDispatch = entry;
  res.json({ hasData: true, uploadedAt: entry.uploadedAt, uploadedBy: entry.uploadedBy, summary: entry.summary, date: entry.date });
});

app.post('/api/dispatch/ask', async (req, res) => {
  try {
    if (!currentDispatch) return res.json({ result: 'No dispatch data loaded. Please upload today\'s report.' });
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001', max_tokens: 1500,
      messages: [{ role: 'user', content: `You are AZHAR-AI Dispatch Intelligence for UAE logistics.\n\nDispatch Summary:\n- Date: ${currentDispatch.date}\n- Total orders: ${currentDispatch.summary.total_orders}\n- Total value: AED ${currentDispatch.summary.total_value?.toLocaleString()}\n- Food orders: ${currentDispatch.summary.food_orders} orders, AED ${currentDispatch.summary.food_value?.toLocaleString()}\n- Non-Food orders: ${currentDispatch.summary.non_food_orders} orders, AED ${currentDispatch.summary.non_food_value?.toLocaleString()}\n- 3PL orders: ${currentDispatch.summary.pl_orders}\n- Total routes: ${currentDispatch.summary.total_routes}\n- Total drivers: ${currentDispatch.summary.total_drivers}\n\nCSV Data:\n${currentDispatch.csvText.substring(0, 10000)}\n\nQuestion: ${req.body.question}\n\nAnswer with exact numbers. Use AED for currency.` }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─────────────────────────────────────────────────────────────
// REJECTION STORE  (persisted to disk — SURVIVES RESTARTS)
// ─────────────────────────────────────────────────────────────
let rejectionData = null;

// Load persisted rejection on startup
const savedRejection = loadJSON(REJECTION_FILE);
if (savedRejection) {
  rejectionData = savedRejection;
  console.log('Loaded rejection data from disk — uploadedAt:', rejectionData.uploadedAt);
}

app.post('/api/rejection/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    const ext = path.extname(req.file.originalname || '').toLowerCase();
    let wb;
    if (ext === '.xlsx' || ext === '.xls') {
      wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    } else {
      return res.status(400).json({ error: 'Please upload .xlsx or .xls file' });
    }

    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    // Parse rejection data
    const ROOT_COL = 'fina- Root Cause';
    const ORG_COL  = 'Organization';
    const DATE_COL = 'D DATE';

    const isRej = (r) => ['R/D','HOLD','RD'].includes(String(r['Status']||'').trim().toUpperCase());
    const isDel = (r) => String(r['Status']||'').toUpperCase().includes('DELIVER');

    const parseDate = (v) => {
      if (!v) return null;
      if (v instanceof Date) return v;
      const d = new Date(v);
      return isNaN(d) ? null : d;
    };

    // Build month/day/org breakdown
    const orgMap = {};
    const monthMap = {};
    let totalRej = 0, totalDel = 0, totalVal = 0;

    rows.forEach(row => {
      const rej = isRej(row), del = isDel(row);
      if (!rej && !del) return;

      const d = parseDate(row[DATE_COL]);
      const mo = d ? d.getMonth() + 1 : null;
      const day = d ? d.getDate() : null;
      const org = String(row[ORG_COL]||'').trim();
      const rootCause = String(row[ROOT_COL]||'').trim();
      const custName = String(row['CUSTOMER NAME']||'').trim();
      const area = String(row['Area']||'').trim();
      const val = parseFloat(row['Value']) || 0;

      if (del) totalDel++;
      if (rej) { totalRej++; totalVal += val; }

      // ORG level
      if (org) {
        if (!orgMap[org]) orgMap[org] = { tDel:0, tRej:0, val:0, reasons:{}, custs:{}, areas:{}, del:[0,0,0,0,0], rej:[0,0,0,0,0] };
        if (del) { orgMap[org].tDel++; if(mo) orgMap[org].del[mo-1]++; }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val += val;
          if(mo) orgMap[org].rej[mo-1]++;
          if (rootCause) orgMap[org].reasons[rootCause] = (orgMap[org].reasons[rootCause]||0) + 1;
          if (custName)  orgMap[org].custs[custName]    = (orgMap[org].custs[custName]||0) + 1;
          if (area)      orgMap[org].areas[area]        = (orgMap[org].areas[area]||0) + 1;
        }
      }

      // Month/day level
      if (mo) {
        if (!monthMap[mo]) monthMap[mo] = { days:new Set(), tDel:0, tRej:0, val:0, reasons:{}, data:{} };
        if (del) monthMap[mo].tDel++;
        if (rej) { monthMap[mo].tRej++; monthMap[mo].val += val; if(rootCause) monthMap[mo].reasons[rootCause]=(monthMap[mo].reasons[rootCause]||0)+1; if(day)monthMap[mo].days.add(day); }
        if(day) {
          if (!monthMap[mo].data[day]) monthMap[mo].data[day] = { tDel:0, tRej:0, val:0, reasons:{}, custs:{}, areas:{} };
          if (del) monthMap[mo].data[day].tDel++;
          if (rej) {
            monthMap[mo].data[day].tRej++; monthMap[mo].data[day].val += val;
            if (rootCause) monthMap[mo].data[day].reasons[rootCause] = (monthMap[mo].data[day].reasons[rootCause]||0)+1;
            if (custName)  monthMap[mo].data[day].custs[custName]    = (monthMap[mo].data[day].custs[custName]||0)+1;
            if (area)      monthMap[mo].data[day].areas[area]        = (monthMap[mo].data[day].areas[area]||0)+1;
          }
        }
      }
    });

    const fmtVal = (v) => v >= 1000000 ? 'AED '+(v/1000000).toFixed(1)+'M' : 'AED '+Math.round(v/1000)+'K';
    const top5 = (obj) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([l,n])=>({l,n}));
    const top5c = (obj) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([n,c])=>({n,c,v:''}));
    const top6a = (obj) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([a,n])=>({a,n}));

    // Build all combined
    const allReasons = {}, allCusts = {}, allAreas = {}, allDel=[0,0,0,0,0], allRej=[0,0,0,0,0];
    Object.values(orgMap).forEach(v => {
      Object.entries(v.reasons).forEach(([k,n])=>allReasons[k]=(allReasons[k]||0)+n);
      Object.entries(v.custs).forEach(([k,n])=>allCusts[k]=(allCusts[k]||0)+n);
      Object.entries(v.areas).forEach(([k,n])=>allAreas[k]=(allAreas[k]||0)+n);
      v.del.forEach((d,i)=>allDel[i]+=d);
      v.rej.forEach((r,i)=>allRej[i]+=r);
    });

    // Serialize month data
    const monthsOut = {};
    Object.entries(monthMap).forEach(([mo, md]) => {
      const days = Array.from(md.days).sort((a,b)=>a-b);
      const dataOut = {};
      Object.entries(md.data).forEach(([day, dd]) => {
        dataOut[day] = {
          tDel: dd.tDel, tRej: dd.tRej, val: fmtVal(dd.val),
          reasons: top5(dd.reasons), custs: top5c(dd.custs), areas: top6a(dd.areas)
        };
      });
      monthsOut[mo] = { days, tDel: md.tDel, tRej: md.tRej, val: fmtVal(md.val), reasons: top5(md.reasons), data: dataOut };
    });

    // Serialize org data
    const orgsOut = { all: { tDel: totalDel, tRej: totalRej, val: fmtVal(totalVal), del: allDel, rej: allRej, reasons: top5(allReasons), custs: top5c(allCusts), areas: top6a(allAreas) } };
    Object.entries(orgMap).forEach(([org, v]) => {
      orgsOut[org] = { tDel: v.tDel, tRej: v.tRej, val: fmtVal(v.val), del: v.del, rej: v.rej, reasons: top5(v.reasons), custs: top5c(v.custs), areas: top6a(v.areas) };
    });

    rejectionData = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: req.body.uploadedBy || 'Admin',
      fileName: req.file.originalname,
      totalOrders: totalRej + totalDel,
      orgs: orgsOut,
      months: monthsOut
    };

    // PERSIST TO DISK — survives server restart
    saveJSON(REJECTION_FILE, rejectionData);
    console.log('Rejection data saved to disk');

    res.json({ success: true, summary: { totalRej, totalDel, fileName: req.file.originalname } });
  } catch (e) {
    console.error('Rejection upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/rejection/status', (req, res) => {
  if (!rejectionData) return res.json({ hasData: false });
  res.json({ hasData: true, uploadedAt: rejectionData.uploadedAt, uploadedBy: rejectionData.uploadedBy, fileName: rejectionData.fileName, totalOrders: rejectionData.totalOrders, orgs: rejectionData.orgs, months: rejectionData.months });
});

// ─── GENERAL CHAT ─────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    let messages = [];
    if (history && history.length > 0) {
      messages = history.slice(-10).map(h => ({ role: h.role === 'assistant' ? 'assistant' : 'user', content: h.content }));
      if (!messages.length || messages[messages.length-1].content !== prompt) messages.push({ role: 'user', content: prompt });
    } else { messages = [{ role: 'user', content: prompt }]; }
    const msg = await anthropic.messages.create({ model: 'claude-haiku-4-5-20251001', max_tokens: 2000, system: 'You are AZHAR-AI, a professional executive assistant for a UAE logistics company. Be concise and helpful.', messages });
    res.json({ result: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
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
      } else { dataText = req.file.buffer.toString('utf8'); }
    }
    const msg = await anthropic.messages.create({ model: 'claude-haiku-4-5-20251001', max_tokens: 2000, messages: [{ role: 'user', content: `${question}\n\n${dataText ? 'Data:\n' + dataText.substring(0, 8000) : ''}` }] });
    res.json({ result: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─── POWERPOINT ───────────────────────────────────────────────
app.post('/api/powerpoint', upload.single('file'), async (req, res) => {
  try {
    const { topic, slides, tone } = req.body;
    let extra = '';
    if (req.file) extra = '\n\nSource content:\n' + req.file.buffer.toString('utf8').substring(0, 3000);
    const msg = await anthropic.messages.create({ model: 'claude-haiku-4-5-20251001', max_tokens: 3000, messages: [{ role: 'user', content: `Create a ${slides || 5}-slide ${tone || 'professional'} presentation about: ${topic}${extra}\n\nFormat each slide as:\nSlide 1: [Title]\n- Bullet point\n\nSpeaker Notes: [notes]` }] });
    res.json({ result: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AZHAR-AI server running on port ${PORT}`));
