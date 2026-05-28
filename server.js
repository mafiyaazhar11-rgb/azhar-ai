const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');

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
    await pool.query(`
      CREATE TABLE IF NOT EXISTS dispatch_data (
        id SERIAL PRIMARY KEY,
        date_key DATE NOT NULL UNIQUE,
        uploaded_at TIMESTAMPTZ DEFAULT NOW(),
        uploaded_by TEXT DEFAULT 'Admin',
        summary JSONB NOT NULL,
        csv_text TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS rejection_data (
        id SERIAL PRIMARY KEY,
        uploaded_at TIMESTAMPTZ DEFAULT NOW(),
        uploaded_by TEXT DEFAULT 'Admin',
        file_name TEXT,
        total_orders INT,
        orgs JSONB,
        months JSONB,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );
    `);
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

app.get('/health', function(req, res) {
  res.json({ status: 'ok', time: new Date().toISOString(), db: !!process.env.DATABASE_URL });
});

//  HELPERS 
function toStr(v) { return String(v == null ? '' : v).trim(); }

function normaliseType(raw) {
  var t = toStr(raw).toUpperCase().replace(/\s+/g, ' ').trim();
  if (t === 'FOOD' || t.startsWith('FOOD')) return 'food';
  if (t === 'NON FOOD' || t.includes('NON FOOD') || t.includes('NON-FOOD')) return 'nonfood';
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
    driver:   findCol('DRIVER CONTACT DETAILS', 'DRIVERS NAME', 'DRIVER NAME', 'DRIVER CONTACT', 'DRIVER_CONTACT'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    type:     findCol('TYPE'),
    org:      findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU') || findCol('ORG')
  };
  console.log('Dispatch cols:', JSON.stringify(C));

  var totalOrders=0, totalValue=0, foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  var cities={}, customers={}, routes={}, driverSet={};
  var orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, HCP:{o:0,v:0} };

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
      if (!routes[route]) routes[route] = { locs:{}, driver:'', value:0 };
      var loc = C.location ? toStr(row[C.location]) : '';
      if (loc) routes[route].locs[loc] = 1;
      routes[route].value += amt;
      if (C.driver && row[C.driver] && !routes[route].driver)
        routes[route].driver = extractDriverName(row[C.driver]);
    }
    if (C.driver && row[C.driver]) {
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
    return { route:route, drops:Object.keys(routes[route].locs).length, driver:routes[route].driver, value:Math.round(routes[route].value) };
  }).sort(function(a,b) { return b.drops-a.drops; }).slice(0,30);

  var driverDrops = {};
  Object.keys(routes).forEach(function(route) {
    var v = routes[route]; if (!v.driver) return;
    driverDrops[v.driver] = (driverDrops[v.driver]||0) + Object.keys(v.locs).length;
  });
  var topDrivers = Object.keys(driverDrops).map(function(name) {
    return { name:name, orders:driverDrops[name] };
  }).sort(function(a,b) { return b.orders-a.orders; }).slice(0,5);

  return {
    total_orders: totalOrders, total_value: Math.round(totalValue),
    total_routes: Object.keys(routes).length,
    total_drivers: Object.keys(driverSet).length || Object.keys(routes).length,
    total_drops: Object.keys(routes).reduce(function(s,r) { return s+Object.keys(routes[r].locs).length; }, 0),
    food_orders: foodOrders, food_value: Math.round(foodValue),
    non_food_orders: nonFoodOrders, non_food_value: Math.round(nonFoodValue),
    pl_orders: plOrders, van_orders: vanOrders,
    type_breakdown: {
      DCV: { orders:orgStats.DCV.o, value:Math.round(orgStats.DCV.v) },
      DCF: { orders:orgStats.DCF.o, value:Math.round(orgStats.DCF.v) },
      DGC: { orders:orgStats.DGC.o, value:Math.round(orgStats.DGC.v) },
      DGS: { orders:orgStats.DGS.o, value:Math.round(orgStats.DGS.v) },
      DSN: { orders:orgStats.DSN.o, value:Math.round(orgStats.DSN.v) },
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
  // Fallback to file
  var saved = loadJSON(DISPATCH_FILE);
  if (saved) {
    dispatchHistory = saved.history || {};
    var keys = Object.keys(dispatchHistory).sort().reverse();
    if (keys.length) currentDispatch = dispatchHistory[keys[0]];
    console.log('Loaded dispatch from file:', keys.length, 'dates');
  }
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
    var keys = Object.keys(dispatchHistory).sort();
    while (keys.length > 60) delete dispatchHistory[keys.shift()];
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
  var saved = loadJSON(REJECTION_FILE);
  if (saved) { rejectionData = saved; console.log('Loaded rejection from file'); }
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
      status: findC('FINAL STATUS', 'STATUS'),
      org:    findC('ORGANIZATION') || findC('ORG-BU'),
      date:   findC('D DATE', 'DATE', 'DELIVERY DATE'),
      root:   findC('FINA- ROOT', 'ROOT CAUSE', 'ROOT_CAUSE', 'REASON-1'),
      cust:   findC('CUSTOMER NAME', 'CUSTOMER'),
      area:   findC('AREA', 'CITY'),
      value:  findC('VALUE', 'AMOUNT'),
      type:   findC('TYPE'),
      source: findC('REMAKE -3', 'REMAKE') || findC('INTERNAL/EXTERNAL')
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
      var val=parseFloat(row[RC.value])||0;
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

var PORT=process.env.PORT||3000;
app.listen(PORT,function(){console.log('AZHAR-AI server running on port '+PORT+(process.env.DATABASE_URL?' with PostgreSQL':' file-only mode'));});
