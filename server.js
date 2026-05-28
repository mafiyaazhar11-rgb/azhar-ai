const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 150 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ extended: true, limit: '150mb' }));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

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

app.get('/health', function(req, res) {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

function toStr(v) { return String(v == null ? '' : v).trim(); }

function normaliseType(raw) {
  var t = toStr(raw).toUpperCase().replace(/\s+/g, ' ');
  if (t === 'FOOD') return 'food';
  if (t === 'NON FOOD') return 'nonfood';
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
  // If it looks like a pure name (no digits), return as-is
  if (/^[A-Za-z][A-Za-z\s]{2,}$/.test(s)) return s.trim();
  // Try to extract name before phone number
  var m = s.match(/^([A-Za-z][A-Za-z\s]{2,29})(?:\s*[-+\d])/);
  if (m) return m[1].trim();
  // Split on digit/dash - but only if result is more than 1 char
  var parts = s.split(/[-+\d]/);
  var name = (parts[0]||'').trim();
  return name.length > 1 ? name : s.trim();
}

function stripBranch(name) {
  var base = toStr(name);
  var kws = [',Branch',', Branch',',Br.',', Br.',' -Branch',',CPD',' CPD','- Branch','-Branch'];
  for (var i = 0; i < kws.length; i++) {
    var idx = base.toLowerCase().indexOf(kws[i].toLowerCase());
    if (idx > 3) { base = base.substring(0, idx).trim(); break; }
  }
  return base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
}

// Find best sheet: most rows with Status + Organization columns
function findDataSheet(wb) {
  var bestSheet = wb.SheetNames[0];
  var bestRows = 0;
  for (var i = 0; i < wb.SheetNames.length; i++) {
    var name = wb.SheetNames[i];
    var ws = wb.Sheets[name];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (rows.length === 0) continue;
    var cols = Object.keys(rows[0]).map(function(c) { return c.toUpperCase(); });
    var hasStatus = cols.some(function(c) { return c.includes('STATUS'); });
    var hasOrg = cols.some(function(c) { return c.includes('ORGAN') || c === 'ORG'; });
    if (hasStatus && hasOrg && rows.length > bestRows) {
      bestRows = rows.length;
      bestSheet = name;
    }
  }
  console.log('Using sheet:', bestSheet, 'with', bestRows, 'rows');
  return bestSheet;
}

function parseDispatch(buffer) {
  var wb = XLSX.read(buffer, { type: 'buffer' });
  var sheetName = findDataSheet(wb);
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!rows.length) return null;

  function findCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
    }) || null;
  }

  var C = {
    route: findCol('ROUTE'),
    city: findCol('CITY', 'AREA'),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount: findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    driver: findCol('DRIVERS NAME', 'DRIVER NAME', 'DRIVER_NAME') || findCol('DRIVER CONTACT', 'DRIVER_CONTACT', 'DRIVER'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    type: findCol('TYPE'),
    org: findCol('ORGANIZATION') || findCol('ORG-BU') || findCol('ORG')
  };
  console.log('Dispatch cols:', JSON.stringify(C));

  var totalOrders=0, totalValue=0, foodOrders=0, foodValue=0;
  var nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  var cities={}, customers={}, routes={}, driverSet={};
  var orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, HCP:{o:0,v:0} };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    totalOrders++;
    var amt = parseFloat(row[C.amount]) || 0;
    totalValue += amt;
    var type = normaliseType(C.type ? row[C.type] : '');
    if (type === 'food') { foodOrders++; foodValue += amt; }
    else if (type === 'nonfood') { nonFoodOrders++; nonFoodValue += amt; }
    else if (type === '3pl') { plOrders++; }
    else if (type === 'van') { vanOrders++; }

    var org = C.org ? toStr(row[C.org]).toUpperCase() : '';
    if (orgStats[org]) { orgStats[org].o++; orgStats[org].v += amt; }
    else if (org === '3 PL' || org === 'HCP') { orgStats.HCP.o++; orgStats.HCP.v += amt; }

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

  console.log('TYPE counts food:'+foodOrders+' nonfood:'+nonFoodOrders+' 3pl:'+plOrders+' van:'+vanOrders+' total:'+totalOrders);

  var byCity = Object.keys(cities).map(function(city) {
    return { city:city, orders:cities[city].orders, value:Math.round(cities[city].value) };
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
    var v = routes[route];
    return { route:route, drops:Object.keys(v.locs).length, driver:v.driver, value:Math.round(v.value) };
  }).sort(function(a,b) { return b.drops-a.drops; }).slice(0,30);

  var driverDrops = {};
  Object.keys(routes).forEach(function(route) {
    var v = routes[route];
    if (!v.driver) return;
    driverDrops[v.driver] = (driverDrops[v.driver]||0) + Object.keys(v.locs).length;
  });
  var topDrivers = Object.keys(driverDrops).map(function(name) {
    return { name:name, orders:driverDrops[name] };
  }).sort(function(a,b) { return b.orders-a.orders; }).slice(0,5);

  return {
    total_orders: totalOrders,
    total_value: Math.round(totalValue),
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
    by_city: byCity,
    top_customers: topCustomers,
    top_drivers: topDrivers,
    top_routes: topRoutes
  };
}

// DISPATCH STORE
var dispatchHistory = {};
var currentDispatch = null;
var savedDispatch = loadJSON(DISPATCH_FILE);
if (savedDispatch) {
  dispatchHistory = savedDispatch.history || {};
  var dkeys = Object.keys(dispatchHistory).sort().reverse();
  if (dkeys.length) currentDispatch = dispatchHistory[dkeys[0]];
  console.log('Loaded dispatch:', dkeys.length, 'dates');
}

app.post('/api/dispatch/upload', upload.single('file'), function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    var summary = parseDispatch(req.file.buffer);
    if (!summary) return res.status(400).json({ error: 'Could not parse file' });
    var dateKey = req.body.dateKey || new Date().toISOString().split('T')[0];
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var wb2 = XLSX.read(req.file.buffer, { type: 'buffer' });
    var csv = XLSX.utils.sheet_to_csv(wb2.Sheets[wb2.SheetNames[0]]);
    var entry = { uploadedAt:new Date().toISOString(), uploadedBy:uploadedBy, csvText:csv.substring(0,200000), summary:summary, date:dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;
    var dkeys2 = Object.keys(dispatchHistory).sort();
    while (dkeys2.length > 30) { delete dispatchHistory[dkeys2.shift()]; }
    saveJSON(DISPATCH_FILE, { history:dispatchHistory });
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
    }).then(function(msg) {
      res.json({ result: msg.content[0].text });
    }).catch(function(e) {
      res.status(500).json({ error: e.message });
    });
  } catch(e) {
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

// REJECTION STORE
var rejectionData = null;
var savedRejection = loadJSON(REJECTION_FILE);
if (savedRejection) {
  rejectionData = savedRejection;
  console.log('Loaded rejection data uploadedAt:', rejectionData.uploadedAt);
}

app.post('/api/rejection/upload', upload.single('file'), function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error:'No file received' });
    var ext = path.extname(req.file.originalname||'').toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls' && ext !== '.csv') return res.status(400).json({ error:'Please upload .xlsx, .xls or .csv' });

    console.log('Reading rejection file:', req.file.originalname, req.file.size, 'bytes');
    var rows = [];
    if (ext === '.csv') {
      // Fast proper CSV parsing (handles quoted fields with commas)
      var csvText = req.file.buffer.toString('utf8');
      var csvRows = csvText.split('\n').filter(function(l){return l.trim();});
      if (csvRows.length < 2) return res.status(400).json({ error:'CSV file is empty' });
      function parseCSVLine(line) {
        var result = [], cell = '', inQ = false;
        for (var ci2=0; ci2<line.length; ci2++) {
          var ch = line[ci2];
          if (ch === '"') { inQ = !inQ; }
          else if (ch === ',' && !inQ) { result.push(cell.trim()); cell = ''; }
          else { cell += ch; }
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
      console.log('Rejection rows:', rows.length, 'in sheet:', sheetName);
    }
    if (!rows.length) return res.status(400).json({ error:'No rows found in file' });

    var keys0 = Object.keys(rows[0]);
    function findC() {
      var names = Array.prototype.slice.call(arguments);
      return keys0.find(function(k) {
        return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
      }) || null;
    }

    var RC = {
      status:    findC('FINAL STATUS', 'STATUS'),
      org:       findC('ORGANIZATION') || findC('ORG-BU'),
      date:      findC('D DATE', 'DATE', 'DELIVERY DATE'),
      root:      findC('FINA- ROOT', 'ROOT CAUSE', 'ROOT_CAUSE', 'REASON-1'),
      cust:      findC('CUSTOMER NAME', 'CUSTOMER'),
      area:      findC('AREA', 'CITY'),
      value:     findC('VALUE', 'AMOUNT'),
      type:      findC('TYPE'),
      source:    findC('REMAKE -3', 'REMAKE') || findC('INTERNAL/EXTERNAL')
    };
    console.log('Rejection cols:', JSON.stringify(RC));

    // Status check - handles both Status and Final Status columns
    function isRej(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1 === 'REJECTION' || s1 === 'R/D' || s1 === 'HOLD' || s1 === 'RD' || s1 === 'REJECTED' || s1 === 'R' ||
             s2 === 'R/D' || s2 === 'HOLD' || s2 === 'REJECTED';
    }
    function isDel(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1 === 'DELIVERED' || s1.includes('DELIVER') || s1 === 'D' ||
             s2.includes('DELIVER') || s2 === 'D' || s2 === 'DELIVERED';
    }

    function parseDate(v) {
      if (!v) return null;
      if (v instanceof Date) return v;
      if (typeof v === 'number') {
        try {
          var unix = Math.round((v - 25569) * 86400 * 1000);
          var dd = new Date(unix);
          if (!isNaN(dd.getTime())) return dd;
        } catch(e2) {}
      }
      var d = new Date(v);
      return isNaN(d.getTime()) ? null : d;
    }

    var orgMap = {}, monthMap = {};
    var totalRej=0, totalDel=0, totalVal=0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rej = isRej(row);
      var del = isDel(row);
      if (!rej && !del) continue;

      var d = parseDate(row[RC.date]);
      var mo  = d ? d.getMonth()+1 : null;
      var day = d ? d.getDate()    : null;
      var org  = toStr(row[RC.org]).toUpperCase().replace('NON-FOOD','DGC');
      var root = toStr(row[RC.root]);
      var cust = toStr(row[RC.cust]);
      var area = toStr(row[RC.area]);
      var val  = parseFloat(row[RC.value]) || 0;

      if (del) totalDel++;
      if (rej) { totalRej++; totalVal += val; }

      if (org) {
        if (!orgMap[org]) orgMap[org] = {
          tDel:0, tRej:0, val:0,
          del: [0,0,0,0,0,0,0,0,0,0,0,0],
          rej: [0,0,0,0,0,0,0,0,0,0,0,0],
          food_rej:0, food_del:0, nonfood_rej:0, nonfood_del:0,
          ext_rej:0, ext_del:0, int_rej:0, int_del:0,
          food_val:0, nonfood_val:0,
          reasons:{}, custs:{}, areas:{}
        };
        if (del) {
          orgMap[org].tDel++; if(mo) orgMap[org].del[mo-1]++;
          var typeStrD = toStr(row[RC.type]||'').toUpperCase();
          var isFoodD = typeStrD === 'FOOD' || typeStrD.startsWith('FOOD,');
          var isNFD = typeStrD.includes('NON FOOD') || typeStrD.includes('NON-FOOD');
          var srcStrD = toStr(row[RC.source]||'').toUpperCase();
          if(isFoodD) orgMap[org].food_del++;
          else if(isNFD) orgMap[org].nonfood_del++;
          if(srcStrD === 'EXTERNAL') orgMap[org].ext_del++;
          else if(srcStrD === 'INTERNAL') orgMap[org].int_del++;
        }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val += val;
          if (mo) orgMap[org].rej[mo-1]++;
          if (root) orgMap[org].reasons[root] = (orgMap[org].reasons[root]||0)+1;
          if (cust) orgMap[org].custs[cust]   = (orgMap[org].custs[cust]  ||0)+1;
          if (area) orgMap[org].areas[area]   = (orgMap[org].areas[area]  ||0)+1;
          // Track type breakdown
          var typeStr = toStr(row[RC.type]||'').toUpperCase();
          var isFood = typeStr === 'FOOD' || typeStr.startsWith('FOOD,');
          var isNonFood = typeStr.includes('NON FOOD') || typeStr.includes('NON-FOOD');
          if (isFood) { orgMap[org].food_rej++; orgMap[org].food_val = (orgMap[org].food_val||0)+val; }
          else if (isNonFood) { orgMap[org].nonfood_rej++; orgMap[org].nonfood_val = (orgMap[org].nonfood_val||0)+val; }
          // Track source breakdown
          var srcStr = toStr(row[RC.source]||'').toUpperCase();
          if (srcStr === 'EXTERNAL') orgMap[org].ext_rej++;
          else if (srcStr === 'INTERNAL') orgMap[org].int_rej++;
        }
      }

      if (mo) {
        if (!monthMap[mo]) monthMap[mo] = { days:{}, tDel:0, tRej:0, val:0, reasons:{}, custs:{}, areas:{}, data:{} };
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++; monthMap[mo].val += val;
          if (root) monthMap[mo].reasons[root] = (monthMap[mo].reasons[root]||0)+1;
          if (cust) monthMap[mo].custs[cust] = (monthMap[mo].custs[cust]||0)+1;
          if (area) monthMap[mo].areas[area] = (monthMap[mo].areas[area]||0)+1;
          if (day) monthMap[mo].days[day] = 1;
        }
        if (day) {
          if (!monthMap[mo].data[day]) monthMap[mo].data[day] = { tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{} };
          if (del) monthMap[mo].data[day].tDel++;
          if (rej) {
            monthMap[mo].data[day].tRej++; monthMap[mo].data[day].val += val;
            if (root) monthMap[mo].data[day].reasons[root] = (monthMap[mo].data[day].reasons[root]||0)+1;
            if (cust) monthMap[mo].data[day].custs[cust]   = (monthMap[mo].data[day].custs[cust]  ||0)+1;
            if (area) monthMap[mo].data[day].areas[area]   = (monthMap[mo].data[day].areas[area]  ||0)+1;
          }
        }
      }
    }

    console.log('Rejection parsed: totalRej='+totalRej+' totalDel='+totalDel);

    function fmtVal(v) { return v>=1000000?'AED '+(v/1000000).toFixed(2)+'M':'AED '+Math.round(v/1000)+'K'; }
    function top10(obj) {
      return Object.keys(obj).map(function(l) { return {l:l,n:obj[l]}; }).sort(function(a,b) { return b.n-a.n; }).slice(0,10);
    }
    function top8c(obj) {
      return Object.keys(obj).map(function(n) { return {n:n,c:obj[n],v:''}; }).sort(function(a,b) { return b.c-a.c; }).slice(0,8);
    }
    function top6a(obj) {
      return Object.keys(obj).map(function(a) { return {a:a,n:obj[a]}; }).sort(function(a,b) { return b.n-a.n; }).slice(0,6);
    }

    var allR={}, allC={}, allA={};
    var allDel = [0,0,0,0,0,0,0,0,0,0,0,0];
    var allRej = [0,0,0,0,0,0,0,0,0,0,0,0];
    Object.keys(orgMap).forEach(function(org2) {
      var v = orgMap[org2];
      Object.keys(v.reasons).forEach(function(k) { allR[k]=(allR[k]||0)+v.reasons[k]; });
      Object.keys(v.custs).forEach(function(k)   { allC[k]=(allC[k]||0)+v.custs[k]; });
      Object.keys(v.areas).forEach(function(k)   { allA[k]=(allA[k]||0)+v.areas[k]; });
      v.del.forEach(function(d,i) { allDel[i]+=d; });
      v.rej.forEach(function(r,i) { allRej[i]+=r; });
    });

    var monthsOut = {};
    Object.keys(monthMap).forEach(function(mo2) {
      var md = monthMap[mo2];
      var dataOut = {};
      Object.keys(md.data).forEach(function(day) {
        var dd = md.data[day];
        dataOut[day] = { tDel:dd.tDel, tRej:dd.tRej, val:fmtVal(dd.val), reasons:top10(dd.reasons), custs:top8c(dd.custs), areas:top6a(dd.areas) };
      });
      monthsOut[mo2] = { days:Object.keys(md.days).map(Number).sort(function(a,b){return a-b;}), tDel:md.tDel, tRej:md.tRej, val:fmtVal(md.val), reasons:top10(md.reasons), custs:top8c(md.custs||{}), areas:top6a(md.areas||{}), data:dataOut };
    });

    var allFoodRej=0,allNonFoodRej=0,allExtRej=0,allIntRej=0,allFoodVal=0,allNonFoodVal=0;
    Object.keys(orgMap).forEach(function(k){var v=orgMap[k];allFoodRej+=v.food_rej||0;allNonFoodRej+=v.nonfood_rej||0;allExtRej+=v.ext_rej||0;allIntRej+=v.int_rej||0;allFoodVal+=v.food_val||0;allNonFoodVal+=v.nonfood_val||0;});
    var orgsOut = { all:{ tDel:totalDel, tRej:totalRej, val:fmtVal(totalVal), food_rej:allFoodRej, nonfood_rej:allNonFoodRej, ext_rej:allExtRej, int_rej:allIntRej, food_val:fmtVal(allFoodVal), nonfood_val:fmtVal(allNonFoodVal), del:allDel, rej:allRej, reasons:top10(allR), custs:top8c(allC), areas:top6a(allA) } };
    Object.keys(orgMap).forEach(function(org3) {
      var v = orgMap[org3];
      orgsOut[org3] = { 
        tDel:v.tDel, tRej:v.tRej, val:fmtVal(v.val), 
        food_rej:v.food_rej||0, food_del:v.food_del||0,
        nonfood_rej:v.nonfood_rej||0, nonfood_del:v.nonfood_del||0,
        ext_rej:v.ext_rej||0, ext_del:v.ext_del||0,
        int_rej:v.int_rej||0, int_del:v.int_del||0,
        food_val:fmtVal(v.food_val||0), nonfood_val:fmtVal(v.nonfood_val||0),
        del:v.del, rej:v.rej, 
        reasons:top10(v.reasons), custs:top8c(v.custs), areas:top6a(v.areas) 
      };
    });

    rejectionData = { uploadedAt:new Date().toISOString(), uploadedBy:req.body.uploadedBy||'Admin', fileName:req.file.originalname, totalOrders:totalRej+totalDel, orgs:orgsOut, months:monthsOut };
    saveJSON(REJECTION_FILE, rejectionData);
    console.log('Rejection saved successfully');

    res.json({ success:true, summary:{ totalRej:totalRej, totalDel:totalDel, fileName:req.file.originalname } });
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
    var prompt = req.body.prompt;
    var history = req.body.history || [];
    var messages = history.slice(-10).map(function(h) {
      return { role: h.role === 'assistant' ? 'assistant' : 'user', content: h.content };
    });
    if (!messages.length || messages[messages.length-1].content !== prompt) {
      messages.push({ role:'user', content:prompt });
    }
    anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:2000,
      system:'You are AZHAR-AI, a professional executive assistant for a UAE logistics company.',
      messages:messages
    }).then(function(msg) {
      res.json({ result: msg.content[0].text });
    }).catch(function(e) {
      res.status(500).json({ error: e.message });
    });
  } catch(e) {
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.post('/api/excel', upload.single('file'), function(req, res) {
  try {
    var question = req.body.question || 'Analyse this data';
    var dataText = '';
    if (req.file) {
      var ext2 = path.extname(req.file.originalname||'').toLowerCase();
      if (ext2 === '.xlsx' || ext2 === '.xls') {
        var wb3 = XLSX.read(req.file.buffer, { type:'buffer' });
        dataText = XLSX.utils.sheet_to_csv(wb3.Sheets[wb3.SheetNames[0]]);
      } else {
        dataText = req.file.buffer.toString('utf8');
      }
    }
    anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:2000,
      messages:[{ role:'user', content:question+(dataText?'\n\nData:\n'+dataText.substring(0,8000):'') }]
    }).then(function(msg) {
      res.json({ result: msg.content[0].text });
    }).catch(function(e) {
      res.status(500).json({ error: e.message });
    });
  } catch(e) {
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

// STATIC - MUST BE LAST
app.get('/', function(req, res) {
  var p1 = path.join(__dirname, 'public', 'index.html');
  var p2 = path.join(__dirname, 'index.html');
  var p3 = path.join(__dirname, 'azhar-ai-v4.html');
  if (fs.existsSync(p1)) return res.sendFile(p1);
  if (fs.existsSync(p2)) return res.sendFile(p2);
  if (fs.existsSync(p3)) return res.sendFile(p3);
  res.status(404).json({ error:'index.html not found' });
});
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

app.use(function(err, req, res, next) {
  console.error('Global error:', err.message);
  if (!res.headersSent) res.status(500).json({ error: err.message || 'Server error' });
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() { console.log('AZHAR-AI server running on port ' + PORT); });
