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

// Serve index.html
app.get('/', (req, res) => {
  const p1 = path.join(__dirname, 'public', 'index.html');
  const p2 = path.join(__dirname, 'index.html');
  if (fs.existsSync(p1)) res.sendFile(p1);
  else if (fs.existsSync(p2)) res.sendFile(p2);
  else res.send('AZHAR-AI Running');
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok' }));


// Direct CSV parser - tuned for AZHAR-AI dispatch format
// Columns: ORDER CODE, ROUTE, Keep Together, CITY, LOCATION_ID, CUSTOMER, 
//          CUSTOMER ADDRESS, TOTAL_AMOUNT, ETA, SPECIAL INSTRUCTIONS, 
//          VEHICLE_ID, DRIVER_ID, DRIVER CONTACT DETAILS

function parseDispatchCSV(csvText, dateKey) {
  const lines = csvText.split('\n').filter(l => l.trim());
  if (lines.length < 2) return { total_orders: 0, date: dateKey };
  
  // Parse headers - find column indices
  const headers = lines[0].split(',').map(h => h.trim().replace(/['"]/g,'').toUpperCase());
  const idx = (names) => {
    for (const n of names) {
      const i = headers.findIndex(h => h.includes(n));
      if (i >= 0) return i;
    }
    return -1;
  };
  
  const routeIdx   = idx(['ROUTE']);
  const cityIdx    = idx(['CITY']);
  const custIdx    = idx(['CUSTOMER']);
  const amtIdx     = idx(['TOTAL_AMOUNT','AMOUNT']);
  const driverIdx   = idx(['DRIVER CONTACT','DRIVER_CONTACT']);
  const driverIdIdx = idx(['DRIVER_ID']);
  const orderIdx    = idx(['ORDER CODE','ORDER_CODE']);
  const locationIdx = idx(['LOCATION_ID','LOCATION']);

  let totalOrders = 0, totalValue = 0;
  const cities = {}, customers = {}, routes = {}, driverSet = new Set();

  for (let i = 1; i < lines.length; i++) {
    // Handle CSV with commas inside quoted fields
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

    // City - normalize case
    if (cityIdx >= 0 && row[cityIdx]) {
      const rawCity = row[cityIdx].trim();
      // Normalize: Title Case, fix variations
      let city = rawCity.charAt(0).toUpperCase() + rawCity.slice(1).toLowerCase();
      if (city.toLowerCase().includes('abu dhabi') || city.toLowerCase() === 'abu dhabi') city = 'Abu Dhabi';
      else if (city.toLowerCase().includes('ras al') || city.toLowerCase().includes('rak')) city = 'Ras Al Khaimah';
      else if (city.toLowerCase().includes('umm al')) city = 'Umm Al Quwain';
      else if (city.toLowerCase() === 'al ain' || city.toLowerCase() === 'al-ain') city = 'Al Ain';
      else if (city.toLowerCase() === 'dubai') city = 'Dubai';
      else if (city.toLowerCase() === 'sharjah') city = 'Sharjah';
      else if (city.toLowerCase() === 'ajman') city = 'Ajman';
      else if (city.toLowerCase() === 'fujairah') city = 'Fujairah';
      else if (city.toLowerCase() === 'hatta') city = 'Hatta';
      if (!cities[city]) cities[city] = { orders: 0, value: 0 };
      cities[city].orders++;
      cities[city].value += amt;
    }

    // Customer
    if (custIdx >= 0 && row[custIdx]) {
      const cust = row[custIdx].trim();
      if (!customers[cust]) customers[cust] = { orders: 0, value: 0 };
      customers[cust].orders++;
      customers[cust].value += amt;
    }

    // Route - drops = unique LOCATION_ID per route
    if (routeIdx >= 0 && row[routeIdx]) {
      const route = row[routeIdx].trim();
      const locId = locationIdx >= 0 ? row[locationIdx]?.trim() : '';
      let driverName = '';
      if (driverIdx >= 0 && row[driverIdx]) {
        const contact = row[driverIdx].trim();
        // Extract name before phone: "Faiz Ullah 056-5362317" -> "Faiz Ullah"
        const match = contact.match(/^([A-Za-z][A-Za-z\s]+?)(?:[\-\+\s]+[0-9]|$)/);
        driverName = match ? match[1].trim() : contact.split(/[\-\+0-9]/)[0].trim();
      }
      if (!routes[route]) routes[route] = { locations: new Set(), driver: driverName, value: 0, orderLines: 0 };
      if (locId) routes[route].locations.add(locId);
      routes[route].orderLines++;
      routes[route].value += amt;
      if (driverName && !routes[route].driver) routes[route].driver = driverName;
    }

    // Driver unique count using DRIVER_ID
    if (driverIdIdx >= 0 && row[driverIdIdx]) driverSet.add(row[driverIdIdx].trim());
    else if (routeIdx >= 0 && row[routeIdx]) driverSet.add(row[routeIdx].trim()); // one driver per route
  }

  // Sort cities
  const byCity = Object.entries(cities)
    .map(([city, v]) => ({ city, orders: v.orders, value: Math.round(v.value) }))
    .sort((a,b) => b.orders - a.orders);

  // Group customers by base name (club same customer across branches/locations)
  const baseCust = {};
  Object.entries(customers).forEach(([name, v]) => {
    // Extract base customer name - remove branch/location specific text
    let base = name;
    const stripAfter = [',Branch', ', Branch', ',Br.', ', Br.', ' -Branch', 
                        ',CPD', ' CPD', '- Branch', '-Branch'];
    for (const kw of stripAfter) {
      const idx = base.toLowerCase().indexOf(kw.toLowerCase());
      if (idx > 3) { base = base.substring(0, idx).trim(); break; }
    }
    // Also strip trailing LLC, L.L.C variations if after comma
    base = base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
    
    if (!baseCust[base]) baseCust[base] = { orders: 0, value: 0, locations: 0 };
    baseCust[base].orders += v.orders;
    baseCust[base].value += v.value;
    baseCust[base].locations++;
  });

  const topCustomers = Object.entries(baseCust)
    .map(([name, v]) => ({ name, orders: v.orders, value: Math.round(v.value), locations: v.locations }))
    .sort((a,b) => b.value - a.value)
    .slice(0, 6);

  // Top routes - drops = unique locations
  const topRoutes = Object.entries(routes)
    .map(([route, v]) => ({ 
      route, 
      drops: v.locations.size,   // unique locations = drops
      order_lines: v.orderLines, // total order lines
      driver: v.driver, 
      value: Math.round(v.value) 
    }))
    .sort((a,b) => b.drops - a.drops)
    .slice(0, 30);

  // Top drivers - from routes sorted by drops
  const topDrivers = Object.entries(routes)
    .filter(([,v]) => v.driver)
    .map(([route, v]) => ({ name: v.driver, orders: v.locations.size, route }))
    .sort((a,b) => b.orders - a.orders)
    .slice(0, 5);

  // Lulu stats
  const luluEntries = Object.entries(customers).filter(([n]) => n.toUpperCase().includes('LULU'));
  const luluOrders = luluEntries.reduce((s,[,v]) => s + v.orders, 0);
  const luluValue  = luluEntries.reduce((s,[,v]) => s + v.value, 0);

  // Food vs Non-food - count by customer type or order prefix
  let foodOrders = 0;
  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(',');
    const order = orderIdx >= 0 ? (row[orderIdx]||'').trim() : '';
    // Food orders typically have DCV or Food type prefix
    if (order.includes('DCV') || order.includes('FOOD')) foodOrders++;
  }

  return {
    total_orders: totalOrders,
    total_value: Math.round(totalValue),
    total_routes: Object.keys(routes).length,
    total_drivers: driverSet.size || Object.keys(routes).length,
    lulu_orders: luluOrders,
    lulu_value: Math.round(luluValue),
    food_orders: foodOrders,
    food_value: Math.round(foodValue),
    non_food_orders: nonFoodOrders,
    non_food_value: Math.round(nonFoodValue),
    pl_orders: plOrders,
    by_city: byCity,
    top_customers: topCustomers,
    top_drivers: topDrivers,
    top_routes: topRoutes,
    date: dateKey
  };
}

// ─── DISPATCH MEMORY STORE ───────────────────────────────────
let dispatchHistory = {};
let currentDispatch = null;

// Upload dispatch
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

    const dateKey = req.body.dateKey || new Date().toISOString().split('T')[0];
    const uploadedBy = req.body.uploadedBy || 'Admin';

    // Generate summary with AI
    // Parse CSV directly for accuracy
    const summary = parseDispatchCSV(csvText, dateKey);
    
    // Top customers already calculated accurately in parseDispatchCSV - no AI needed

    const entry = { uploadedAt: new Date().toISOString(), uploadedBy, csvText, summary, date: dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;

    // Keep only 30 days
    const keys = Object.keys(dispatchHistory).sort();
    if (keys.length > 30) delete dispatchHistory[keys[0]];

    res.json({ success: true, summary, uploadedAt: entry.uploadedAt, date: dateKey });
  } catch (e) {
    console.error('Upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Get status
app.get('/api/dispatch/status', (req, res) => {
  const availableDates = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData: false, availableDates });
  res.json({
    hasData: true,
    uploadedAt: currentDispatch.uploadedAt,
    uploadedBy: currentDispatch.uploadedBy,
    summary: currentDispatch.summary,
    date: currentDispatch.date,
    availableDates
  });
});

// Load specific date
app.get('/api/dispatch/date/:dateKey', (req, res) => {
  const entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData: false });
  currentDispatch = entry;
  res.json({ hasData: true, uploadedAt: entry.uploadedAt, uploadedBy: entry.uploadedBy, summary: entry.summary, date: entry.date });
});

// Ask dispatch question
app.post('/api/dispatch/ask', async (req, res) => {
  try {
    if (!currentDispatch) return res.json({ result: 'No dispatch data loaded. Please ask admin to upload today\'s report.' });
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      messages: [{
        role: 'user',
        content: `You are AZHAR-AI Dispatch Intelligence for UAE logistics company in UAE.

IMPORTANT FACTS about this dispatch data:
- Total orders: ${currentDispatch.summary.total_orders || 'unknown'}
- Total value: AED ${currentDispatch.summary.total_value || 'unknown'}
- Food orders (numeric order codes): ${currentDispatch.summary.food_orders || 'unknown'} orders, AED ${currentDispatch.summary.food_value || 'unknown'}
- Non-food/3PL orders (HCP codes): ${currentDispatch.summary.non_food_orders || 'unknown'} orders
- Total routes: ${currentDispatch.summary.total_routes || 'unknown'}
- Total drivers: ${currentDispatch.summary.total_drivers || 'unknown'}
- Date: ${currentDispatch.date}

CSV Data (for detailed queries):
${currentDispatch.csvText.substring(0, 10000)}

Columns: ORDER CODE, ROUTE, CITY, LOCATION_ID, CUSTOMER, TOTAL_AMOUNT, ETA, VEHICLE_ID, DRIVER_ID, DRIVER CONTACT DETAILS

Note: Numeric order codes = Food orders. HCP_ prefix = Non-food/3PL orders.

Question: ${req.body.question}

Answer with exact numbers from the data above. Use AED for currency values.`
      }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── GENERAL CHAT ─────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    let messages = [];
    if (history && history.length > 0) {
      messages = history.slice(-10).map(h => ({ role: h.role === 'assistant' ? 'assistant' : 'user', content: h.content }));
      if (!messages.length || messages[messages.length-1].content !== prompt) {
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
      messages: [{ role: 'user', content: `${question}\n\n${dataText ? 'Data:\n' + dataText.substring(0, 8000) : ''}` }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
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
      messages: [{ role: 'user', content: `Create a ${slides || 5}-slide ${tone || 'professional'} presentation about: ${topic}${extra}\n\nFormat each slide as:\nSlide 1: [Title]\n- Bullet point\n- Bullet point\n\nSpeaker Notes: [notes]` }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AZHAR-AI server running on port ${PORT}`));
