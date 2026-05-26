const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
// Create uploads dir if not exists
if (!fs.existsSync('uploads')) { fs.mkdirSync('uploads', { recursive: true }); }

// Use memory storage for reliability on cloud
const storage = multer.memoryStorage();
const upload = multer({ storage: storage, limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));
app.use(express.static('.'));
app.get('/', (req, res) => {
  const path = require('path');
  const fs = require('fs');
  // Try public/index.html first, then root index.html
  const publicPath = path.join(__dirname, 'public', 'index.html');
  const rootPath = path.join(__dirname, 'index.html');
  if (fs.existsSync(publicPath)) {
    res.sendFile(publicPath);
  } else if (fs.existsSync(rootPath)) {
    res.sendFile(rootPath);
  } else {
    res.send('AZHAR-AI Server is running! index.html not found.');
  }
});

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Health check
app.get('/health', (req, res) => res.json({ status: 'AZHAR-AI server running' }));

// In-memory dispatch data store - keeps 30 days history
let dispatchHistory = {}; // keyed by date string YYYY-MM-DD
let dispatchData = {
  uploadedAt: null,
  uploadedBy: null,
  csvText: null,
  summary: null,
  date: null
};

// Upload dispatch data (admin only)
app.post('/api/dispatch/upload', upload.single('file'), async (req, res) => {
  try {
    let csvText = '';
    if (req.file) {
      const ext = path.extname(req.file.originalname).toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        // Handle Excel file from buffer
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        csvText = XLSX.utils.sheet_to_csv(sheet);
        console.log('Excel converted from buffer, size:', csvText.length);
      } else {
        // Handle CSV/text file from buffer
        csvText = req.file.buffer.toString('utf8');
        console.log('CSV from buffer, size:', csvText.length, 'chars');
      }
    } else if (req.body.csvText) {
      csvText = req.body.csvText;
      console.log('CSV text received, size:', csvText.length, 'chars');
    }
    if (!csvText) return res.status(400).json({ error: 'No data provided' });

    // Generate auto summary using Claude
    const summaryPrompt = `You are AZHAR-AI Dispatch Intelligence. Analyse this dispatch CSV data and provide a structured JSON summary.

CSV Data (first 12000 chars):
${csvText.substring(0, 12000)}

Return ONLY a JSON object (no markdown, no explanation) with this exact structure:
{
  "total_orders": <number>,
  "total_value": <number>,
  "total_routes": <number>,
  "total_drivers": <number>,
  "by_city": [{"city": "Dubai", "orders": 0, "value": 0}],
  "top_customers": [{"name": "Customer", "orders": 0, "value": 0}],
  "top_drivers": [{"name": "Driver", "orders": 0}],
  "lulu_orders": <number>,
  "lulu_value": <number>,
  "food_orders": <number>,
  "non_food_orders": <number>,
  "date": "<date from data>"
}
Return ONLY the JSON. food_orders = count of orders where CUSTOMER name contains FOOD. non_food_orders = total_orders minus food_orders.`;

    const message = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1000,
      messages: [{ role: 'user', content: summaryPrompt }]
    });

    let summary = null;
    try {
      const raw = message.content[0].text.replace(/```json|```/g, '').trim();
      summary = JSON.parse(raw);
    } catch(e) {
      summary = { error: 'Could not parse summary' };
    }

    dispatchData = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: req.body.uploadedBy || 'Admin',
      csvText: csvText,
      summary
    };

    res.json({ success: true, summary, uploadedAt: dispatchData.uploadedAt });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Get dispatch status + available dates
app.get('/api/dispatch/status', (req, res) => {
  const availableDates = Object.keys(dispatchHistory).sort().reverse();
  if (!dispatchData.uploadedAt) {
    return res.json({ hasData: false, availableDates });
  }
  res.json({
    hasData: true,
    uploadedAt: dispatchData.uploadedAt,
    uploadedBy: dispatchData.uploadedBy,
    summary: dispatchData.summary,
    date: dispatchData.date,
    availableDates
  });
});

// Load specific date
app.get('/api/dispatch/date/:dateKey', (req, res) => {
  const { dateKey } = req.params;
  const entry = dispatchHistory[dateKey];
  if (!entry) {
    return res.json({ hasData: false, message: 'No data for this date' });
  }
  // Set as current active data
  dispatchData = entry;
  res.json({
    hasData: true,
    uploadedAt: entry.uploadedAt,
    uploadedBy: entry.uploadedBy,
    summary: entry.summary,
    date: entry.date
  });
});

// Ask dispatch question
app.post('/api/dispatch/ask', async (req, res) => {
  try {
    if (!dispatchData.csvText) {
      return res.json({ result: 'No dispatch data loaded yet. Please ask your admin to upload today dispatch report.' });
    }
    const { question } = req.body;
    const prompt = `You are AZHAR-AI Dispatch Intelligence for a logistics/distribution company in UAE.

Today's dispatch data (CSV):
${dispatchData.csvText.substring(0, 12000)}

Data uploaded: ${dispatchData.uploadedAt}

The columns are: ORDER CODE, ROUTE, CITY, CUSTOMER, CUSTOMER ADDRESS, TOTAL_AMOUNT, ETA, VEHICLE_ID, DRIVER_ID, DRIVER CONTACT DETAILS

User question: ${question}

Answer accurately based on the data. Be specific with numbers, names and values. Format nicely with bullet points where helpful. Always mention AED for currency values. If asking about Lulu - search for customers containing "LULU" or "Lulu" in the name.`;

    const message = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 1500,
      messages: [{ role: 'user', content: prompt }]
    });
    res.json({ result: message.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Text AI call with history support
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, history } = req.body;
    let messages = [];
    // Add history if provided
    if (history && history.length > 0) {
      messages = history.slice(-10).map(h => ({
        role: h.role === 'assistant' ? 'assistant' : 'user',
        content: h.content
      }));
      // Make sure last message is the current prompt
      if (messages[messages.length-1]?.content !== prompt) {
        messages.push({ role: 'user', content: prompt });
      }
    } else {
      messages = [{ role: 'user', content: prompt }];
    }
    const message = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 2000,
      system: 'You are AZHAR-AI, a highly professional and intelligent executive assistant. You help with any question, task or request. Be concise, accurate and professional.',
      messages: messages
    });
    res.json({ result: message.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Excel file upload + analyse
app.post('/api/excel', upload.single('file'), async (req, res) => {
  try {
    const { question } = req.body;
    let fileContent = '';

    if (req.file) {
      const ext = path.extname(req.file.originalname).toLowerCase();
      if (ext === '.xlsx' || ext === '.xls') {
        const workbook = XLSX.readFile(req.file.path);
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          fileContent += `\nSheet: ${sheetName}\n`;
          fileContent += XLSX.utils.sheet_to_csv(sheet);
        });
      } else {
        fileContent = fs.readFileSync(req.file.path, 'utf8');
      }
      fs.unlinkSync(req.file.path);
    }

    const prompt = `You are AZHAR-AI, an expert Excel analyst.
${fileContent ? `Excel file data:\n${fileContent.substring(0, 8000)}` : ''}
${question ? `User request: ${question}` : 'Analyse this data, find errors, fix formulas, and give a summary.'}

Provide:
1. Direct answer to the request
2. Key findings from the data
3. Any formula fixes needed
4. Recommendations`;

    const message = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 2000,
      messages: [{ role: 'user', content: prompt }]
    });
    res.json({ result: message.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// PowerPoint generate + return slide data
app.post('/api/powerpoint', upload.single('file'), async (req, res) => {
  try {
    const { topic, slides, tone } = req.body;
    let fileContent = '';

    if (req.file) {
      const ext = path.extname(req.file.originalname).toLowerCase();
      if (ext === '.pptx' || ext === '.docx') {
        const result = await mammoth.extractRawText({ path: req.file.path });
        fileContent = result.value;
      } else {
        fileContent = fs.readFileSync(req.file.path, 'utf8');
      }
      fs.unlinkSync(req.file.path);
    }

    const prompt = `You are AZHAR-AI, a professional presentation builder.
Create a ${slides || 8}-slide PowerPoint presentation.
Topic: ${topic || 'Based on the uploaded content'}
Tone: ${tone || 'Professional'}
${fileContent ? `Source content:\n${fileContent.substring(0, 6000)}` : ''}

Format EXACTLY like this:
Slide 1: [Title]
- [bullet]
- [bullet]
- [bullet]
Speaker Note: [note]

Slide 2: [Title]
...and so on for all ${slides || 8} slides.`;

    const message = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 3000,
      messages: [{ role: 'user', content: prompt }]
    });
    res.json({ result: message.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`AZHAR-AI server running on port ${PORT}`));
