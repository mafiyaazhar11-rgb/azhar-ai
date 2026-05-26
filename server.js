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
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: `Analyse this dispatch CSV and return ONLY a JSON object, no markdown, no explanation:
{
  "total_orders": <count all rows>,
  "total_value": <sum of TOTAL_AMOUNT column>,
  "total_routes": <count unique ROUTE values>,
  "total_drivers": <count unique DRIVER_ID values>,
  "lulu_orders": <count rows where CUSTOMER contains LULU>,
  "lulu_value": <sum TOTAL_AMOUNT where CUSTOMER contains LULU>,
  "food_orders": <count rows where CUSTOMER contains FOOD>,
  "non_food_orders": <total_orders minus food_orders>,
  "by_city": [{"city":"Dubai","orders":0,"value":0}],
  "top_customers": [{"name":"Customer","orders":0,"value":0}],
  "top_drivers": [{"name":"Driver","orders":0}],
  "date": "${dateKey}"
}
top_customers: top 6 by value. top_drivers: top 5 by order count.

CSV (first 10000 chars):
${csvText.substring(0, 10000)}`
      }]
    });

    let summary = {};
    try {
      const raw = msg.content[0].text.replace(/```json|```/g, '').trim();
      summary = JSON.parse(raw);
    } catch(e) {
      summary = { total_orders: 0, error: 'Could not parse summary' };
    }

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
        content: `You are AZHAR-AI Dispatch Intelligence for UAE logistics.
Dispatch CSV data:
${currentDispatch.csvText.substring(0, 12000)}

Columns: ORDER CODE, ROUTE, CITY, CUSTOMER, TOTAL_AMOUNT, ETA, VEHICLE_ID, DRIVER_ID, DRIVER CONTACT DETAILS
Date: ${currentDispatch.date}

Question: ${req.body.question}

Answer accurately with specific numbers. Use AED for currency.`
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
