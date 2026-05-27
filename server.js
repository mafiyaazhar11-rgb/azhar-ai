const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('.'));

app.get('/', (req, res) => {
  const p = path.join(__dirname, 'index.html');
  if (fs.existsSync(p)) res.sendFile(p);
  else res.send('AZHAR-AI Running');
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ── DISPATCH STORE ─────────────────────────────────────────────
let dispatchHistory = {};
let currentDispatch = null;

function parseDispatchCSV(csvText, dateKey) {
  const lines = csvText.split('\n').filter(l => l.trim());
  if (lines.length < 2) return { total_orders: 0, date: dateKey };

  const headers = lines[0].split(',').map(h => h.trim().replace(/['"]/g,'').toUpperCase());
  const col = (names) => { for (const n of names) { const i = headers.findIndex(h => h.includes(n)); if (i >= 0) return i; } return -1; };

  const routeIdx    = col(['ROUTE']);
  const cityIdx     = col(['CITY']);
  const custIdx     = col(['CUSTOMER']);
  const amtIdx      = col(['TOTAL_AMOUNT','AMOUNT']);
  const driverIdx   = col(['DRIVER CONTACT','DRIVER_CONTACT']);
  const driverIdIdx = col(['DRIVER_ID']);
  const orderIdx    = col(['ORDER CODE','ORDER_CODE']);
  const locationIdx = col(['LOCATION_ID','LOCATION']);
  const typeColIdx  = col(['TYPE']);

  let totalOrders=0, totalValue=0;
  const cities={}, customers={}, routes={}, driverSet=new Set();
  const typeStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, HCP:{o:0,v:0} };
  let foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;

  for (let i = 1; i < lines.length; i++) {
    const row = [];
    let cell = '', inQ = false;
    for (const ch of lines[i]) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { row.push(cell.trim()); cell = ''; }
      else { cell += ch; }
    }
    row.push(cell.trim());
    if (!row[0]) continue;
    totalOrders++;

    const amt = amtIdx >= 0 ? parseFloat(row[amtIdx]) || 0 : 0;
    totalValue += amt;

    // City
    if (cityIdx >= 0 && row[cityIdx]) {
      let city = row[cityIdx].trim();
      const cu = city.toUpperCase();
      if (cu.includes('ABU DHABI') || cu === 'ABUDHABI') city = 'Abu Dhabi';
      else if (cu.includes('RAS AL') || cu === 'RAK') city = 'Ras Al Khaimah';
      else if (cu.includes('UMM AL')) city = 'Umm Al Quwain';
      else if (cu === 'AL AIN' || cu === 'AL-AIN') city = 'Al Ain';
      else city = city.charAt(0).toUpperCase() + city.slice(1).toLowerCase();
      if (!cities[city]) cities[city] = { orders:0, value:0 };
      cities[city].orders++; cities[city].value += amt;
    }

    // Customer
    if (custIdx >= 0 && row[custIdx]) {
      const cust = row[custIdx].trim();
      if (!customers[cust]) customers[cust] = { orders:0, value:0 };
      customers[cust].orders++; customers[cust].value += amt;
    }

    // Route + drops
    if (routeIdx >= 0 && row[routeIdx]) {
      const route = row[routeIdx].trim();
      const locId = locationIdx >= 0 ? (row[locationIdx]||'').trim() : '';
      let driverName = '';
      if (driverIdx >= 0 && row[driverIdx]) {
        const contact = row[driverIdx].trim();
        const m = contact.match(/^([A-Za-z][A-Za-z\s]+?)(?:[\-\+\s]+[0-9]|$)/);
        driverName = m ? m[1].trim() : contact.split(/[\-\+0-9]/)[0].trim();
      }
      if (!routes[route]) routes[route] = { locations: new Set(), driver: driverName, value:0, orderLines:0 };
      if (locId) routes[route].locations.add(locId);
      routes[route].orderLines++;
      routes[route].value += amt;
      if (driverName && !routes[route].driver) routes[route].driver = driverName;
    }

    // Driver count
    if (driverIdIdx >= 0 && row[driverIdIdx]) driverSet.add(row[driverIdIdx].trim());
    else if (routeIdx >= 0 && row[routeIdx]) driverSet.add(row[routeIdx].trim());

    // Type classification
    if (typeColIdx >= 0) {
      const tv = (row[typeColIdx]||'').trim().toUpperCase();
      const isFood    = tv === 'FOOD';
      const isNonFood = tv.includes('NON') || tv.includes('NOON');
      const is3PL     = tv.includes('3') && tv.includes('P');
      const isVan     = tv.includes('VAN');
      if      (isFood)    { foodOrders++;    foodValue    += amt; typeStats.DCV.o++; typeStats.DCV.v += amt; }
      else if (isNonFood) { nonFoodOrders++; nonFoodValue += amt; typeStats.DGC.o++; typeStats.DGC.v += amt; }
      else if (is3PL)     { plOrders++;  typeStats.HCP.o++; }
      else if (isVan)     { vanOrders++; }
    } else {
      const order = (orderIdx >= 0 ? row[orderIdx] : '').trim().toUpperCase();
      const m = order.match(/[A-Z]{2,3}/);
      const tc = m ? m[0] : '';
      if (typeStats[tc]) { typeStats[tc].o++; typeStats[tc].v += amt; }
      if      (tc==='DCV'||tc==='DCF') { foodOrders++;    foodValue    += amt; }
      else if (tc==='DGC'||tc==='DGS'||tc==='DSN') { nonFoodOrders++; nonFoodValue += amt; }
      else if (tc==='HCP') plOrders++;
    }
  }

  // City list sorted
  const byCity = Object.entries(cities)
    .map(([city,v]) => ({ city, orders:v.orders, value:Math.round(v.value) }))
    .sort((a,b) => b.orders - a.orders);

  // Group customers by base name
  const baseCust = {};
  Object.entries(customers).forEach(([name,v]) => {
    let base = name;
    for (const kw of [',Branch',', Branch',',Br.',', Br.',' -Branch','-Branch']) {
      const idx = base.toLowerCase().indexOf(kw.toLowerCase());
      if (idx > 3) { base = base.substring(0, idx).trim(); break; }
    }
    base = base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i,'').trim();
    if (!baseCust[base]) baseCust[base] = { orders:0, value:0, locations:0 };
    baseCust[base].orders   += v.orders;
    baseCust[base].value    += v.value;
    baseCust[base].locations++;
  });

  const topCustomers = Object.entries(baseCust)
    .map(([name,v]) => ({ name, orders:v.orders, value:Math.round(v.value), locations:v.locations }))
    .sort((a,b) => b.value - a.value).slice(0,6);

  const topRoutes = Object.entries(routes)
    .map(([route,v]) => ({ route, drops:v.locations.size, driver:v.driver, value:Math.round(v.value), order_lines:v.orderLines }))
    .sort((a,b) => b.drops - a.drops).slice(0,30);

  const topDrivers = Object.entries(routes)
    .filter(([,v]) => v.driver)
    .map(([,v]) => ({ name:v.driver, orders:v.locations.size }))
    .sort((a,b) => b.orders - a.orders).slice(0,5);

  const luluEntries = Object.entries(customers).filter(([n]) => n.toUpperCase().includes('LULU'));

  return {
    total_orders:    totalOrders,
    total_value:     Math.round(totalValue),
    total_routes:    Object.keys(routes).length,
    total_drivers:   driverSet.size || Object.keys(routes).length,
    total_drops:     Object.values(routes).reduce((s,r) => s + r.locations.size, 0),
    food_orders:     foodOrders,
    food_value:      Math.round(foodValue),
    non_food_orders: nonFoodOrders,
    non_food_value:  Math.round(nonFoodValue),
    pl_orders:       plOrders,
    van_orders:      vanOrders,
    lulu_orders:     luluEntries.reduce((s,[,v]) => s+v.orders, 0),
    lulu_value:      Math.round(luluEntries.reduce((s,[,v]) => s+v.value, 0)),
    type_breakdown: {
      DCV: { orders:typeStats.DCV.o, value:Math.round(typeStats.DCV.v) },
      DCF: { orders:typeStats.DCF.o, value:Math.round(typeStats.DCF.v) },
      DGC: { orders:typeStats.DGC.o, value:Math.round(typeStats.DGC.v) },
      DGS: { orders:typeStats.DGS.o, value:Math.round(typeStats.DGS.v) },
      DSN: { orders:typeStats.DSN.o, value:Math.round(typeStats.DSN.v) },
      HCP: { orders:typeStats.HCP.o, value:Math.round(typeStats.HCP.v) }
    },
    by_city:       byCity,
    top_customers: topCustomers,
    top_drivers:   topDrivers,
    top_routes:    topRoutes,
    date:          dateKey
  };
}

// ── DISPATCH ENDPOINTS ─────────────────────────────────────────
app.post('/api/dispatch/upload', upload.single('file'), async (req, res) => {
  try {
    let csvText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname||'').toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        const wb = XLSX.read(req.file.buffer, { type:'buffer' });
        csvText = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
      } else {
        csvText = req.file.buffer.toString('utf8');
      }
    } else if (req.body.csvText) {
      csvText = req.body.csvText;
    }
    if (!csvText) return res.status(400).json({ error:'No file received' });

    const dateKey    = req.body.dateKey || new Date().toISOString().split('T')[0];
    const uploadedBy = req.body.uploadedBy || 'Admin';
    const summary    = parseDispatchCSV(csvText, dateKey);
    const entry      = { uploadedAt: new Date().toISOString(), uploadedBy, csvText, summary, date: dateKey };

    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;

    const keys = Object.keys(dispatchHistory).sort();
    if (keys.length > 30) delete dispatchHistory[keys[0]];

    res.json({ success:true, summary, uploadedAt:entry.uploadedAt, date:dateKey });
  } catch(e) { console.error(e); res.status(500).json({ error:e.message }); }
});

app.get('/api/dispatch/status', (req, res) => {
  const availableDates = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData:false, availableDates });
  res.json({ hasData:true, uploadedAt:currentDispatch.uploadedAt, uploadedBy:currentDispatch.uploadedBy, summary:currentDispatch.summary, date:currentDispatch.date, availableDates });
});

app.get('/api/dispatch/date/:dateKey', (req, res) => {
  const entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData:false });
  currentDispatch = entry;
  res.json({ hasData:true, uploadedAt:entry.uploadedAt, uploadedBy:entry.uploadedBy, summary:entry.summary, date:entry.date });
});

app.post('/api/dispatch/ask', async (req, res) => {
  try {
    if (!currentDispatch) return res.json({ result:'No dispatch data loaded. Please upload today\'s report.' });
    const s = currentDispatch.summary;
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001', max_tokens: 1500,
      messages: [{ role:'user', content:
        `AZHAR-AI Dispatch Intelligence — UAE Logistics\n\nKEY FACTS:\n- Total orders: ${s.total_orders}\n- Total value: AED ${s.total_value}\n- Food orders: ${s.food_orders} (AED ${s.food_value})\n- Non-Food orders: ${s.non_food_orders} (AED ${s.non_food_value})\n- 3PL orders: ${s.pl_orders}\n- Routes: ${s.total_routes}, Drivers: ${s.total_drivers}, Drops: ${s.total_drops}\n- Date: ${s.date}\n\nCSV DATA:\n${currentDispatch.csvText.substring(0,10000)}\n\nQuestion: ${req.body.question}\n\nAnswer with exact numbers. Use AED for currency.`
      }]
    });
    res.json({ result: msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

// ── OTHER ENDPOINTS ────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    const messages = history && history.length
      ? [...history.slice(-10).map(h => ({ role: h.role==='assistant'?'assistant':'user', content:h.content })), { role:'user', content:prompt }]
      : [{ role:'user', content:prompt }];
    const msg = await anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:2000,
      system:'You are AZHAR-AI, a professional executive assistant for a UAE logistics company.',
      messages
    });
    res.json({ result: msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

app.post('/api/excel', upload.single('file'), async (req, res) => {
  try {
    let question = req.body.question || 'Analyse this data';
    let dataText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname||'').toLowerCase();
      if (ext==='.xlsx'||ext==='.xls') {
        const wb = XLSX.read(req.file.buffer, { type:'buffer' });
        dataText = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
      } else { dataText = req.file.buffer.toString('utf8'); }
    }
    const msg = await anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:2000,
      messages:[{ role:'user', content: question + (dataText ? '\n\nData:\n' + dataText.substring(0,8000) : '') }]
    });
    res.json({ result: msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AZHAR-AI server running on port ${PORT}`));
