const express = require('express');
const multer = require('multer');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/', limits: { fileSize: 50 * 1024 * 1024 } });

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

// Text AI call
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt } = req.body;
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
