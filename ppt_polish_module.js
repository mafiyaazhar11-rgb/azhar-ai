// ============================================================
// PPT POLISH MODULE — "upload a rough PowerPoint, get it polished"
// Pure Node, no Python/LibreOffice needed. A .pptx is a zip of XML;
// images live as independent parts, completely separate from the text.
// This never touches the image bytes — it only reads text out, asks
// Claude to improve it, and rebuilds a new deck with the SAME images
// placed into the new design.
//
// v2: auto-detects the brand's accent color from the deck's own logo,
// gives section-divider slides a distinct styled treatment, and will
// generate a real native chart for a slide's numbers — but ONLY when
// Claude explicitly says it's confident in the label-to-number pairing.
// Anything less than confident falls back to safe bullets. A wrong
// chart in front of a CEO is worse than a plain one — this errs toward
// never showing a number next to the wrong label.
//
// Mount with: require('./ppt_polish_module')(app, requireAuth, upload, anthropic, auditLog);
// Reuses the existing `anthropic` client already configured in server.js —
// this uses your existing Claude API key and existing billing, same as
// Email Writer and the other AI features already in this app.
// ============================================================

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');
const { PNG } = require('pngjs');

module.exports = function (app, requireAuth, upload, anthropic, auditLog) {

  const DEFAULT_ACCENT = 'C9A84C'; // fallback gold, used whenever color detection isn't possible

  // ---------- extraction (same core as v1) ----------
  function extractTexts(parsedSlideXml) {
    const results = [];
    function walk(obj) {
      if (obj === null || typeof obj !== 'object') return;
      if (obj['a:t'] !== undefined) {
        const t = obj['a:t'];
        const text = typeof t === 'string' ? t : (t && t['#text']) || '';
        if (text.trim()) results.push(text.trim());
      }
      for (const key in obj) {
        const val = obj[key];
        if (Array.isArray(val)) val.forEach(walk);
        else if (typeof val === 'object') walk(val);
      }
    }
    walk(parsedSlideXml);
    return results;
  }
  function extractEmbeddedRelIds(parsedSlideXml) {
    const ids = [];
    function walk(obj) {
      if (obj === null || typeof obj !== 'object') return;
      if (obj['a:blip'] !== undefined) {
        const blip = obj['a:blip'];
        const embed = blip && (blip['@_r:embed'] || blip['@_r:link']);
        if (embed) ids.push(embed);
      }
      for (const key in obj) {
        const val = obj[key];
        if (Array.isArray(val)) val.forEach(walk);
        else if (typeof val === 'object') walk(val);
      }
    }
    walk(parsedSlideXml);
    return ids;
  }
  function getSlideOrder(zip) {
    const parser = new XMLParser({ ignoreAttributes: false });
    const presXml = parser.parse(zip.readAsText('ppt/presentation.xml'));
    const relsXml = parser.parse(zip.readAsText('ppt/_rels/presentation.xml.rels'));
    const relMap = {};
    let rels = relsXml.Relationships.Relationship;
    if (!Array.isArray(rels)) rels = [rels];
    rels.forEach(r => { relMap[r['@_Id']] = r['@_Target']; });
    let sldIds = presXml['p:presentation']['p:sldIdLst']['p:sldId'];
    if (!Array.isArray(sldIds)) sldIds = [sldIds];
    return sldIds.map(s => 'ppt/' + relMap[s['@_r:id']].replace(/^\.?\//, ''));
  }

  async function extractDeck(buffer) {
    const zip = new AdmZip(buffer);
    const parser = new XMLParser({ ignoreAttributes: false });
    const slidePaths = getSlideOrder(zip);
    const slides = [];
    for (const slidePath of slidePaths) {
      try {
        const parsed = parser.parse(zip.readAsText(slidePath));
        const texts = extractTexts(parsed);
        const embeddedRelIds = extractEmbeddedRelIds(parsed);
        const slideName = slidePath.split('/').pop();
        const relsPath = `ppt/slides/_rels/${slideName}.rels`;
        const images = [];
        const relsEntry = zip.getEntry(relsPath);
        if (relsEntry && embeddedRelIds.length) {
          const relsParsed = parser.parse(zip.readAsText(relsPath));
          let rels = relsParsed.Relationships.Relationship;
          if (!Array.isArray(rels)) rels = [rels];
          embeddedRelIds.forEach(rid => {
            const rel = rels.find(r => r['@_Id'] === rid);
            if (rel && rel['@_Target'] && rel['@_Target'].includes('media/')) {
              const mediaPath = 'ppt' + rel['@_Target'].replace('..', '');
              const mediaEntry = zip.getEntry(mediaPath);
              if (mediaEntry) images.push({ buffer: mediaEntry.getData(), ext: mediaPath.split('.').pop() });
            }
          });
        }
        slides.push({ texts, images });
      } catch (slideErr) {
        console.error(`PPT polish: could not fully read ${slidePath}:`, slideErr.message);
        slides.push({ texts: [], images: [] });
      }
    }
    return slides;
  }

  // ---------- auto brand-color detection ----------
  // Only handles PNG (pure JS, no native bindings — safe to deploy on
  // Render exactly as-is). JPEG logos fall back to the default gold
  // rather than risk a heavier native image library for this.
  function detectAccentColor(slides) {
    try {
      let candidate = null;
      if (slides[0] && slides[0].images[0] && slides[0].images[0].ext === 'png') {
        candidate = slides[0].images[0];
      } else {
        const freq = {};
        slides.forEach(s => s.images.forEach(img => {
          if (img.ext !== 'png') return;
          const key = img.buffer.length;
          freq[key] = (freq[key] || 0) + 1;
        }));
        const mostCommonSize = Object.entries(freq).sort((a, b) => b[1] - a[1])[0];
        if (mostCommonSize && mostCommonSize[1] > 1) {
          for (const s of slides) {
            const found = s.images.find(img => img.ext === 'png' && img.buffer.length === Number(mostCommonSize[0]));
            if (found) { candidate = found; break; }
          }
        }
      }
      if (!candidate) return DEFAULT_ACCENT;

      const png = PNG.sync.read(candidate.buffer);
      const counts = {};
      for (let i = 0; i < png.data.length; i += 4 * 3) {
        const r = png.data[i], g = png.data[i + 1], b = png.data[i + 2], a = png.data[i + 3];
        if (a < 200) continue;
        if (r > 240 && g > 240 && b > 240) continue;
        if (r < 20 && g < 20 && b < 20) continue;
        const key = [Math.round(r / 16) * 16, Math.round(g / 16) * 16, Math.round(b / 16) * 16].join(',');
        counts[key] = (counts[key] || 0) + 1;
      }
      const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      if (!sorted.length) return DEFAULT_ACCENT;
      const [r, g, b] = sorted[0][0].split(',').map(Number);
      return [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    } catch (e) {
      console.error('PPT polish: color detection failed, using default.', e.message);
      return DEFAULT_ACCENT;
    }
  }

  app.post('/api/ppt-polish', requireAuth, upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      if (!req.file.originalname.toLowerCase().endsWith('.pptx')) {
        return res.status(400).json({ error: 'Only .pptx files are supported (not .ppt — save as .pptx first)' });
      }

      const slides = await extractDeck(req.file.buffer);
      if (!slides.length) return res.status(400).json({ error: 'Could not read any slides from this file' });

      const ACCENT = detectAccentColor(slides);

      const outline = slides.map((s, i) => ({ slide: i + 1, original_text: s.texts, has_images: s.images.length }));
      const prompt = `You are improving a rough, unfinished PowerPoint deck for a professional business/CEO-level presentation.

Here is the raw extracted text per slide (in order), plus how many images each slide has:
${JSON.stringify(outline, null, 2)}

For EACH slide, return a JSON object with:
- "type": "divider" if this slide is clearly just a section header (very short text, like a number + a few words, no real content), otherwise "content"
- "title": an improved, professional title
- "bullets": 2-5 clean bullet points (empty array for divider slides), based on the original content — fix grammar, tighten wording. Preserve factual meaning; never invent facts not implied by the original text.
- "chart": ONLY include this field if the original text contains a clear, unambiguous set of labels each paired with a number (e.g. "Category X: 45%", "Region Y: 1200 units") that can be safely turned into a bar chart. Set "confident": true only if you are certain which number belongs to which label — if there is ANY ambiguity, doubt, or the numbers appear to be from a table where structure may have been lost, omit "chart" entirely or set "confident": false. It is far better to skip a chart than to risk pairing a number with the wrong label. If included: {"confident": true, "chartLabel": "...", "categories": ["...","..."], "values": [number, number, ...]} — categories and values arrays must be the same length.

Keep exactly ${slides.length} slides, in the same order.

Respond with ONLY valid JSON, no markdown fences, no commentary, in this exact shape:
{"slides":[{"type":"content","title":"...","bullets":["...","..."],"chart":null}]}`;

      let msg;
      try {
        msg = await anthropic.messages.create({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 16000,
          messages: [{ role: 'user', content: prompt }],
        });
      } catch (apiErr) {
        console.error('PPT polish: the writing-assistant call itself failed.');
        console.error('  status:', apiErr.status, '| type:', apiErr.error && apiErr.error.type, '| message:', apiErr.message);
        return res.status(500).json({ error: 'Could not reach the writing assistant right now. Please try again in a moment.' });
      }

      let raw = msg.content[0].text.trim();
      raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '');
      const firstBrace = raw.indexOf('{');
      const lastBrace = raw.lastIndexOf('}');
      if (firstBrace !== -1 && lastBrace !== -1) raw = raw.substring(firstBrace, lastBrace + 1);

      let improved;
      try { improved = JSON.parse(raw); } catch (e) {
        console.error('PPT polish: could not parse response as JSON. First 500 chars:', raw.substring(0, 500));
        return res.status(500).json({ error: 'Could not process that file right now. Try again, or try a smaller file.' });
      }
      if (!improved.slides || !Array.isArray(improved.slides) || improved.slides.length === 0) {
        console.error('PPT polish: response had no usable slides array.');
        return res.status(500).json({ error: 'Something went wrong preparing that file. Try again.' });
      }
      while (improved.slides.length < slides.length) {
        const idx = improved.slides.length;
        improved.slides.push({ type: 'content', title: (slides[idx].texts[0] || `Slide ${idx + 1}`), bullets: slides[idx].texts.slice(1, 6) });
      }
      if (improved.slides.length > slides.length) improved.slides = improved.slides.slice(0, slides.length);

      // Defense in depth: never trust the "confident" flag alone. A chart
      // only renders if the data is ALSO structurally sound — matching
      // array lengths, all-numeric values, at least 2 categories.
      function isChartSafe(chart) {
        if (!chart || chart.confident !== true) return false;
        if (!Array.isArray(chart.categories) || !Array.isArray(chart.values)) return false;
        if (chart.categories.length < 2 || chart.categories.length !== chart.values.length) return false;
        return chart.values.every(v => typeof v === 'number' && isFinite(v));
      }

      const pptxgen = require('pptxgenjs');
      const pres = new pptxgen();
      pres.layout = 'LAYOUT_WIDE';
      const INK = '1A1A1A', MUTED = '6B7280';
      const TITLE_FONT = 'Georgia', BODY_FONT = 'Calibri';
      const GREEN = '3CB6AE', RED = 'D9705F'; // chart up/down colors — neutral, not tied to any specific brand palette

      improved.slides.forEach((sl, i) => {
        const s = pres.addSlide();
        const hasImage = slides[i].images.length > 0;
        const chartSafe = isChartSafe(sl.chart);

        if (sl.type === 'divider') {
          s.background = { color: INK };
          s.addShape(pres.ShapeType.rect, { x: 0.9, y: 4.55, w: 0.8, h: 0.06, fill: { color: ACCENT } });
          s.addText(sl.title || `Section ${i + 1}`, { x: 0.9, y: 3.6, w: 10.5, h: 0.9, fontFace: TITLE_FONT, fontSize: 34, color: 'FFFFFF' });
          s.addText(`${i + 1} / ${improved.slides.length}`, { x: 12.4, y: 7.1, w: 0.7, h: 0.3, fontFace: BODY_FONT, fontSize: 9, color: '9CA3AF', align: 'right' });
          return;
        }

        s.background = { color: 'FFFFFF' };
        s.addText(sl.title || `Slide ${i + 1}`, { x: 0.6, y: 0.5, w: 11.8, h: 0.9, fontFace: TITLE_FONT, fontSize: 26, color: INK });

        if (chartSafe) {
          const colors = sl.chart.values.map(v => v >= 0 ? GREEN : RED);
          s.addChart(pres.ChartType.bar, [{ name: sl.chart.chartLabel || '', labels: sl.chart.categories, values: sl.chart.values }], {
            x: 0.5, y: 1.8, w: 12.3, h: 4.4, barDir: 'bar', chartColors: colors,
            showTitle: false, showLegend: false, showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: INK, dataLabelFontSize: 11,
            catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLineShow: false,
            valAxisLabelColor: MUTED, valAxisLabelFontSize: 9, valAxisLineShow: false, valAxisHidden: true,
            valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
          });
        } else {
          const bulletItems = (sl.bullets || []).map((b, bi) => ({ text: b, options: { bullet: { code: '2022', color: ACCENT }, breakLine: bi < sl.bullets.length - 1, color: MUTED, paraSpaceAfter: 10 } }));
          const textW = hasImage ? 6.3 : 11.8;
          if (bulletItems.length) s.addText(bulletItems, { x: 0.6, y: 1.7, w: textW, h: 5.0, fontFace: BODY_FONT, fontSize: 15, valign: 'top' });

          if (hasImage) {
            const img = slides[i].images[0];
            const mime = img.ext === 'jpg' || img.ext === 'jpeg' || img.ext === 'jfif' ? 'image/jpeg' : img.ext === 'gif' ? 'image/gif' : 'image/png';
            try {
              s.addImage({ data: `${mime};base64,${img.buffer.toString('base64')}`, x: 7.2, y: 1.7, w: 5.5, h: 4.5, sizing: { type: 'contain', w: 5.5, h: 4.5 } });
            } catch (e) { /* skip a malformed image rather than fail the whole deck */ }
          }
        }

        s.addText(`${i + 1} / ${improved.slides.length}`, { x: 12.4, y: 7.1, w: 0.7, h: 0.3, fontFace: BODY_FONT, fontSize: 9, color: MUTED, align: 'right' });
      });

      const buf = await pres.write({ outputType: 'nodebuffer' });
      await auditLog(req.user.uid, req.user.username, 'ppt_polished', `${req.file.originalname} -> ${improved.slides.length} slides, accent ${ACCENT}`, req.ip);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="Polished_${req.file.originalname.replace(/\.pptx$/i, '')}.pptx"`);
      res.send(buf);
    } catch (e) {
      console.error('PPT polish error:', e.message);
      res.status(500).json({ error: 'Could not polish that file right now. Please try again.' });
    }
  });

  console.log('PPT Polish module: routes mounted');
};
