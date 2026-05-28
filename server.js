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

// ── DATA DIR ──────────────────────────────────────────────────
const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DISPATCH_FILE  = path.join(DATA_DIR, 'dispatch.json');
const REJECTION_FILE = path.join(DATA_DIR, 'rejection.json');

function saveJSON(fp, data) {
  try { fs.writeFileSync(fp, JSON.stringify(data)); } catch(e) { console.error('Save error:', e.message); }
}
function loadJSON(fp) {
  try { if (fs.existsSync(fp)) return JSON.parse(fs.readFileSync(fp, 'utf8')); } catch(e) {}
  return null;
}

// ── HEALTH ────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// ── HELPERS ───────────────────────────────────────────────────
function toStr(v) { return String(v == null ? '' : v).trim(); }

function normaliseType(raw) {
  const t = toStr(raw).toUpperCase().replace(/\s+/g, ' ');
  if (t === 'FOOD')              return 'food';
  if (t === 'NON FOOD')          return 'nonfood';
  if (t === '3PL' || t === '3 PL') return '3pl';
  if (t === 'VAN')               return 'van';
  return t.toLowerCase();
}

function normaliseCity(raw) {
  const c = toStr(raw).toLowerCase();
  if (c.includes('abu dhabi'))                         return 'Abu Dhabi';
  if (c.includes('dubai'))                             return 'Dubai';
  if (c.includes('sharjah'))                           return 'Sharjah';
  if (c.includes('ajman'))                             return 'Ajman';
  if (c.includes('fujairah'))                          return 'Fujairah';
  if (c.includes('al ain') || c.includes('al-ain'))   return 'Al Ain';
  if (c.includes('ras al') || c === 'rak')             return 'Ras Al Khaimah';
  if (c.includes('umm'))                               return 'Umm Al Quwain';
  const s = toStr(raw);
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

function extractDriverName(contact) {
  const s = toStr(contact);
  if (!s) return '';
  const m = s.match(/^([A-Za-z][A-Za-z\s]{1,29})(?:\s*[-+\d])/);
  if (m) return m[1].trim();
  return s.split(/[-+\d]/)[0].trim();
}

function stripBranch(name) {
  let base = toStr(name);
  const kws = [',Branch',', Branch',',Br.',', Br.',' -Branch',',CPD',' CPD','- Branch','-Branch'];
  for (const kw of kws) {
    const i = base.toLowerCase().indexOf(kw.toLowerCase());
    if (i > 3) { base = base.substring(0, i).trim(); break; }
  }
  return base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
}

// ── DISPATCH PARSER ───────────────────────────────────────────
function parseDispatch(buffer) {
  const wb   = XLSX.read(buffer, { type: 'buffer' });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!rows.length) return null;

  const findCol = (...names) => Object.keys(rows[0]).find(k =>
    names.some(n => k.toUpperCase().includes(n.toUpperCase()))
  ) || null;

  const C = {
    route:    findCol('ROUTE'),
    city:     findCol('CITY'),
    customer: findCol('CUSTOMER'),
    amount:   findCol('TOTAL_AMOUNT','AMOUNT','VALUE'),
    driver:   findCol('DRIVER CONTACT','DRIVER_CONTACT','DRIVER'),
    location: findCol('LOCATION_ID','LOCATION'),
    type:     findCol('TYPE'),
    org:      findCol('ORG'),
  };

  console.log('Dispatch columns detected:', C);

  let totalOrders=0, totalValue=0;
  let foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  const cities={}, customers={}, routes={}, driverSet=new Set();
  const orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, HCP:{o:0,v:0} };

  for (const row of rows) {
    totalOrders++;
    const amt = parseFloat(row[C.amount]) || 0;
    totalValue += amt;

    const type = normaliseType(C.type ? row[C.type] : '');
    if      (type === 'food')    { foodOrders++;    foodValue    += amt; }
    else if (type === 'nonfood') { nonFoodOrders++; nonFoodValue += amt; }
    else if (type === '3pl')     { plOrders++; }
    else if (type === 'van')     { vanOrders++; }

    const org = C.org ? toStr(row[C.org]).toUpperCase() : '';
    if      (org === 'DCV') { orgStats.DCV.o++; orgStats.DCV.v += amt; }
    else if (org === 'DCF') { orgStats.DCF.o++; orgStats.DCF.v += amt; }
    else if (org === 'DGC') { orgStats.DGC.o++; orgStats.DGC.v += amt; }
    else if (org === 'DGS') { orgStats.DGS.o++; orgStats.DGS.v += amt; }
    else if (org === 'DSN') { orgStats.DSN.o++; orgStats.DSN.v += amt; }
    else if (org === '3 PL' || org === 'HCP') { orgStats.HCP.o++; orgStats.HCP.v += amt; }

    if (C.city && row[C.city]) {
      const city = normaliseCity(row[C.city]);
      if (!cities[city]) cities[city] = { orders:0, value:0 };
      cities[city].orders++; cities[city].value += amt;
    }

    if (C.customer && row[C.customer]) {
      const cust = toStr(row[C.customer]);
      if (!customers[cust]) customers[cust] = { orders:0, value:0 };
      customers[cust].orders++; customers[cust].value += amt;
    }

    if (C.route && row[C.route]) {
      const route = toStr(row[C.route]);
      if (!routes[route]) routes[route] = { locations: new Set(), driver:'', value:0 };
      const loc = C.location ? toStr(row[C.location]) : '';
      if (loc) routes[route].locations.add(loc);
      routes[route].value += amt;
      if (C.driver && row[C.driver] && !routes[route].driver)
        routes[route].driver = extractDriverName(row[C.driver]);
    }

    if (C.driver && row[C.driver]) driverSet.add(extractDriverName(row[C.driver]) || toStr(row[C.driver]));
  }

  console.log('TYPE counts — food:'+foodOrders+' nonfood:'+nonFoodOrders+' 3pl:'+plOrders+' van:'+vanOrders+' total:'+totalOrders);

  const byCity = Object.entries(cities)
    .map(([city,v]) => ({ city, orders:v.orders, value:Math.round(v.value) }))
    .sort((a,b) => b.orders - a.orders);

  const baseCust = {};
  for (const [name,v] of Object.entries(customers)) {
    const base = stripBranch(name);
    if (!baseCust[base]) baseCust[base] = { orders:0, value:0 };
    baseCust[base].orders += v.orders;
    baseCust[base].value  += v.value;
  }
  const topCustomers = Object.entries(baseCust)
    .map(([name,v]) => ({ name, orders:v.orders, value:Math.round(v.value) }))
    .sort((a,b) => b.value - a.value).slice(0,6);

  const topRoutes = Object.entries(routes)
    .map(([route,v]) => ({ route, drops:v.locations.size, driver:v.driver, value:Math.round(v.value) }))
    .sort((a,b) => b.drops - a.drops).slice(0,30);

  const driverDrops = {};
  for (const [,v] of Object.entries(routes)) {
    if (!v.driver) continue;
    driverDrops[v.driver] = (driverDrops[v.driver]||0) + v.locations.size;
  }
  const topDrivers = Object.entries(driverDrops)
    .map(([name,orders]) => ({ name, orders }))
    .sort((a,b) => b.orders - a.orders).slice(0,5);

  return {
    total_orders:    totalOrders,
    total_value:     Math.round(totalValue),
    total_routes:    Object.keys(routes).length,
    total_drivers:   driverSet.size || Object.keys(routes).length,
    total_drops:     Object.values(routes).reduce((s,r) => s+r.locations.size, 0),
    food_orders:     foodOrders,
    food_value:      Math.round(foodValue),
    non_food_orders: nonFoodOrders,
    non_food_value:  Math.round(nonFoodValue),
    pl_orders:       plOrders,
    van_orders:      vanOrders,
    type_breakdown: {
      DCV: { orders:orgStats.DCV.o, value:Math.round(orgStats.DCV.v) },
      DCF: { orders:orgStats.DCF.o, value:Math.round(orgStats.DCF.v) },
      DGC: { orders:orgStats.DGC.o, value:Math.round(orgStats.DGC.v) },
      DGS: { orders:orgStats.DGS.o, value:Math.round(orgStats.DGS.v) },
      DSN: { orders:orgStats.DSN.o, value:Math.round(orgStats.DSN.v) },
      HCP: { orders:orgStats.HCP.o, value:Math.round(orgStats.HCP.v) }
    },
    by_city:       byCity,
    top_customers: topCustomers,
    top_drivers:   topDrivers,
    top_routes:    topRoutes
  };
}

// ── DISPATCH STORE ────────────────────────────────────────────
let dispatchHistory = {};
let currentDispatch = null;

const savedDispatch = loadJSON(DISPATCH_FILE);
if (savedDispatch) {
  dispatchHistory = savedDispatch.history || {};
  const keys = Object.keys(dispatchHistory).sort().reverse();
  if (keys.length) currentDispatch = dispatchHistory[keys[0]];
  console.log('Loaded dispatch:', keys.length, 'dates');
}

app.post('/api/dispatch/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    const summary = parseDispatch(req.file.buffer);
    if (!summary) return res.status(400).json({ error: 'Could not parse file' });

    const dateKey    = req.body.dateKey    || new Date().toISOString().split('T')[0];
    const uploadedBy = req.body.uploadedBy || 'Admin';

    const wb  = XLSX.read(req.file.buffer, { type:'buffer' });
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);

    const entry = { uploadedAt:new Date().toISOString(), uploadedBy, csvText:csv.substring(0,200000), summary, date:dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;

    const keys = Object.keys(dispatchHistory).sort();
    while (keys.length > 30) delete dispatchHistory[keys.shift()];
    saveJSON(DISPATCH_FILE, { history:dispatchHistory });

    res.json({ success:true, summary, uploadedAt:entry.uploadedAt, date:dateKey });
  } catch(e) {
    console.error('Dispatch upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
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
    if (!currentDispatch) return res.json({ result:'No dispatch data. Please upload first.' });
    const s = currentDispatch.summary;
    const context = 'Date: '+currentDispatch.date+'\nTotal Orders: '+s.total_orders+'\nTotal Value: AED '+s.total_value+'\nFood: '+s.food_orders+' orders AED '+s.food_value+'\nNon-Food: '+s.non_food_orders+' orders AED '+s.non_food_value+'\n3PL: '+s.pl_orders+'\n\nCSV:\n'+(currentDispatch.csvText||'').substring(0,8000);
    const msg = await anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:1500,
      messages:[{ role:'user', content:'You are AZHAR-AI Dispatch Intelligence for UAE logistics.\n\n'+context+'\n\nQuestion: '+req.body.question+'\n\nAnswer with exact numbers. Use AED for currency.' }]
    });
    res.json({ result: msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

// ── REJECTION STORE ───────────────────────────────────────────
let rejectionData = null;

const savedRejection = loadJSON(REJECTION_FILE);
if (savedRejection) {
  rejectionData = savedRejection;
  console.log('Loaded rejection data — uploadedAt:', rejectionData.uploadedAt);
}

app.post('/api/rejection/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error:'No file received' });
    const ext = path.extname(req.file.originalname||'').toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls') return res.status(400).json({ error:'Please upload .xlsx or .xls' });

    const wb   = XLSX.read(req.file.buffer, { type:'buffer' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval:'' });
    if (!rows.length) return res.status(400).json({ error:'No rows found' });

    const keys0 = Object.keys(rows[0]);
    const findC = (...names) => keys0.find(k => names.some(n => k.toUpperCase().includes(n.toUpperCase()))) || null;

    const RC = {
      status: findC('STATUS'),
      org:    findC('ORGANIZATION','ORG'),
      date:   findC('D DATE','DATE','DELIVERY DATE'),
      root:   findC('ROOT CAUSE','ROOT_CAUSE','FINA'),
      cust:   findC('CUSTOMER NAME','CUSTOMER'),
      area:   findC('AREA','CITY'),
      value:  findC('VALUE','AMOUNT'),
    };
    console.log('Rejection columns:', RC);

    const isRej = r => { const s=toStr(r[RC.status]).toUpperCase(); return s==='R/D'||s==='HOLD'||s==='RD'||s==='REJECTED'||s==='R'; };
    const isDel = r => { const s=toStr(r[RC.status]).toUpperCase(); return s.includes('DELIVER')||s==='D'||s==='DELIVERED'; };

    const parseDate = v => {
      if (!v) return null;
      if (v instanceof Date) return v;
      if (typeof v === 'number') { const d=XLSX.SSF.parse_date_code(v); if(d) return new Date(d.y,d.m-1,d.d); }
      const d = new Date(v); return isNaN(d)?null:d;
    };

    const orgMap={}, monthMap={};
    let totalRej=0, totalDel=0, totalVal=0;

    for (const row of rows) {
      const rej=isRej(row), del=isDel(row);
      if (!rej && !del) continue;

      const d    = parseDate(row[RC.date]);
      const mo   = d ? d.getMonth()+1 : null;
      const day  = d ? d.getDate()    : null;
      const org  = toStr(row[RC.org]).toUpperCase();
      const root = toStr(row[RC.root]);
      const cust = toStr(row[RC.cust]);
      const area = toStr(row[RC.area]);
      const val  = parseFloat(row[RC.value])||0;

      if (del) totalDel++;
      if (rej) { totalRej++; totalVal+=val; }

      if (org) {
        if (!orgMap[org]) orgMap[org]={ tDel:0,tRej:0,val:0, del:new Array(12).fill(0), rej:new Array(12).fill(0), reasons:{},custs:{},areas:{} };
        if (del) { orgMap[org].tDel++; if(mo) orgMap[org].del[mo-1]++; }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val+=val;
          if (mo) orgMap[org].rej[mo-1]++;
          if (root) orgMap[org].reasons[root]=(orgMap[org].reasons[root]||0)+1;
          if (cust) orgMap[org].custs[cust]  =(orgMap[org].custs[cust]  ||0)+1;
          if (area) orgMap[org].areas[area]  =(orgMap[org].areas[area]  ||0)+1;
        }
      }

      if (mo) {
        if (!monthMap[mo]) monthMap[mo]={ days:new Set(),tDel:0,tRej:0,val:0,reasons:{},data:{} };
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++; monthMap[mo].val+=val;
          if (root) monthMap[mo].reasons[root]=(monthMap[mo].reasons[root]||0)+1;
          if (day)  monthMap[mo].days.add(day);
        }
        if (day) {
          if (!monthMap[mo].data[day]) monthMap[mo].data[day]={ tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{} };
          if (del) monthMap[mo].data[day].tDel++;
          if (rej) {
            monthMap[mo].data[day].tRej++; monthMap[mo].data[day].val+=val;
            if (root) monthMap[mo].data[day].reasons[root]=(monthMap[mo].data[day].reasons[root]||0)+1;
            if (cust) monthMap[mo].data[day].custs[cust]  =(monthMap[mo].data[day].custs[cust]  ||0)+1;
            if (area) monthMap[mo].data[day].areas[area]  =(monthMap[mo].data[day].areas[area]  ||0)+1;
          }
        }
      }
    }

    const fmtVal = v => v>=1000000?'AED '+(v/1000000).toFixed(2)+'M':'AED '+Math.round(v/1000)+'K';
    const top10  = obj => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([l,n])=>({l,n}));
    const top8c  = obj => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([n,c])=>({n,c,v:''}));
    const top6a  = obj => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([a,n])=>({a,n}));

    const allR={},allC={},allA={},allDel=new Array(12).fill(0),allRej=new Array(12).fill(0);
    for (const v of Object.values(orgMap)) {
      for (const [k,n] of Object.entries(v.reasons)) allR[k]=(allR[k]||0)+n;
      for (const [k,n] of Object.entries(v.custs))   allC[k]=(allC[k]||0)+n;
      for (const [k,n] of Object.entries(v.areas))   allA[k]=(allA[k]||0)+n;
      v.del.forEach((d,i)=>allDel[i]+=d);
      v.rej.forEach((r,i)=>allRej[i]+=r);
    }

    const monthsOut={};
    for (const [mo,md] of Object.entries(monthMap)) {
      const dataOut={};
      for (const [day,dd] of Object.entries(md.data)) {
        dataOut[day]={ tDel:dd.tDel,tRej:dd.tRej,val:fmtVal(dd.val), reasons:top10(dd.reasons),custs:top8c(dd.custs),areas:top6a(dd.areas) };
      }
      monthsOut[mo]={ days:Array.from(md.days).sort((a,b)=>a-b), tDel:md.tDel,tRej:md.tRej,val:fmtVal(md.val), reasons:top10(md.reasons),data:dataOut };
    }

    const orgsOut={ all:{ tDel:totalDel,tRej:totalRej,val:fmtVal(totalVal), del:allDel,rej:allRej, reasons:top10(allR),custs:top8c(allC),areas:top6a(allA) } };
    for (const [org,v] of Object.entries(orgMap)) {
      orgsOut[org]={ tDel:v.tDel,tRej:v.tRej,val:fmtVal(v.val), del:v.del,rej:v.rej, reasons:top10(v.reasons),custs:top8c(v.custs),areas:top6a(v.areas) };
    }

    rejectionData={ uploadedAt:new Date().toISOString(), uploadedBy:req.body.uploadedBy||'Admin', fileName:req.file.originalname, totalOrders:totalRej+totalDel, orgs:orgsOut, months:monthsOut };
    saveJSON(REJECTION_FILE, rejectionData);
    console.log('Rejection saved:', totalRej, 'rejections,', totalDel, 'delivered');

    res.json({ success:true, summary:{ totalRej,totalDel,fileName:req.file.originalname } });
  } catch(e) {
    console.error('Rejection upload error:', e.message, e.stack);
    res.status(500).json({ error:e.message });
  }
});

// !! CRITICAL — always returns JSON, never HTML !!
app.get('/api/rejection/status', (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  if (!rejectionData) return res.json({ hasData:false });
  res.json({ hasData:true, uploadedAt:rejectionData.uploadedAt, uploadedBy:rejectionData.uploadedBy, fileName:rejectionData.fileName, totalOrders:rejectionData.totalOrders, orgs:rejectionData.orgs, months:rejectionData.months });
});

// ── CHAT ──────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    let messages = history && history.length
      ? [...history.slice(-10).map(h=>({ role:h.role==='assistant'?'assistant':'user', content:h.content })), { role:'user', content:prompt }]
      : [{ role:'user', content:prompt }];
    const msg = await anthropic.messages.create({ model:'claude-haiku-4-5-20251001', max_tokens:2000, system:'You are AZHAR-AI, a professional executive assistant for a UAE logistics company. Be concise and helpful.', messages });
    res.json({ result:msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

// ── EXCEL ─────────────────────────────────────────────────────
app.post('/api/excel', upload.single('file'), async (req, res) => {
  try {
    let question = req.body.question || 'Analyse this data';
    let dataText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname||'').toLowerCase();
      dataText = (ext==='.xlsx'||ext==='.xls')
        ? XLSX.utils.sheet_to_csv(XLSX.read(req.file.buffer,{type:'buffer'}).Sheets[XLSX.read(req.file.buffer,{type:'buffer'}).SheetNames[0]])
        : req.file.buffer.toString('utf8');
    }
    const msg = await anthropic.messages.create({ model:'claude-haiku-4-5-20251001', max_tokens:2000, messages:[{ role:'user', content:question+(dataText?'\n\nData:\n'+dataText.substring(0,8000):'') }] });
    res.json({ result:msg.content[0].text });
  } catch(e) { res.status(500).json({ error:e.message }); }
});

// ── STATIC — MUST BE LAST ─────────────────────────────────────
app.get('/', (req, res) => {
  const p1 = path.join(__dirname, 'public', 'index.html');
  const p2 = path.join(__dirname, 'index.html');
  const p3 = path.join(__dirname, 'azhar-ai-v4.html');
  if (fs.existsSync(p1)) return res.sendFile(p1);
  if (fs.existsSync(p2)) return res.sendFile(p2);
  if (fs.existsSync(p3)) return res.sendFile(p3);
  res.status(404).json({ error:'index.html not found' });
});
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// ── START ─────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('AZHAR-AI server running on port ' + PORT));
