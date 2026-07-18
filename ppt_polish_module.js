// ============================================================
// PPT POLISH MODULE — "upload a rough PowerPoint, get it polished"
// Pure Node, no Python/LibreOffice needed. A .pptx is a zip of XML;
// images live as independent parts, completely separate from the text.
// This never touches the image bytes — it only reads text out, asks
// Claude to improve it, and rebuilds a new deck with the SAME images
// placed into the new design.
//
// Mount with: require('./ppt_polish_module')(app, requireAuth, upload, anthropic, auditLog);
// Reuses the existing `anthropic` client already configured in server.js —
// this uses your existing Claude API key and existing billing, same as
// Email Writer and the other AI features already in this app.
// ============================================================

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

module.exports = function (app, requireAuth, upload, anthropic, auditLog) {

  // Walk parsed slide XML and collect every <a:t> text run in document order.
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

  // Which media relationship IDs does this slide actually embed (<a:blip r:embed="rIdN">)?
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
      const slideXml = zip.readAsText(slidePath);
      const parsed = parser.parse(slideXml);
      const texts = extractTexts(parsed);
      const embeddedRelIds = extractEmbeddedRelIds(parsed);

      // Resolve embedded rel IDs -> actual media file bytes via this slide's .rels
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
    }
    return slides;
  }

  app.post('/api/ppt-polish', requireAuth, upload.single('file'), async function (req, res) {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      if (!req.file.originalname.toLowerCase().endsWith('.pptx')) {
        return res.status(400).json({ error: 'Only .pptx files are supported (not .ppt — save as .pptx first)' });
      }

      const slides = await extractDeck(req.file.buffer);
      if (!slides.length) return res.status(400).json({ error: 'Could not read any slides from this file' });

      // Ask Claude to improve the wording/structure — same slide count and
      // order preserved, since that's what keeps each slide's images aligned
      // with the right new content.
      const outline = slides.map((s, i) => ({ slide: i + 1, original_text: s.texts, has_images: s.images.length }));
      const prompt = `You are improving the text content of a rough, unfinished PowerPoint deck for a professional business presentation.

Here is the raw extracted text per slide (in order), plus how many images each slide has:
${JSON.stringify(outline, null, 2)}

For EACH slide, produce an improved, professional title and 2-5 clean bullet points, based on the original content — fix grammar, tighten wording, make it presentation-ready. Preserve the factual meaning; do not invent facts not implied by the original text. Keep exactly ${slides.length} slides, in the same order, so slide N here corresponds to original slide N.

Respond with ONLY valid JSON, no markdown fences, no commentary, in this exact shape:
{"slides":[{"title":"...","bullets":["...","..."]}]}`;

      const msg = await anthropic.messages.create({
        model: 'claude-sonnet-5',
        max_tokens: 4000,
        messages: [{ role: 'user', content: prompt }],
      });

      let raw = msg.content[0].text.trim();
      raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '');
      let improved;
      try { improved = JSON.parse(raw); } catch (e) {
        return res.status(500).json({ error: 'Could not process that file right now. Try again.' });
      }
      if (!improved.slides || improved.slides.length !== slides.length) {
        return res.status(500).json({ error: 'Something went wrong preparing that file. Try again.' });
      }

      // Rebuild with pptxgenjs — same clean design language as the rest of
      // this app, with the ORIGINAL images (untouched bytes) placed back in.
      const pptxgen = require('pptxgenjs');
      const pres = new pptxgen();
      pres.layout = 'LAYOUT_WIDE';
      const ACCENT = 'C9A84C', INK = '1A1A1A', MUTED = '6B7280';
      const TITLE_FONT = 'Georgia', BODY_FONT = 'Calibri';

      improved.slides.forEach((sl, i) => {
        const s = pres.addSlide();
        s.background = { color: 'FFFFFF' };
        const hasImage = slides[i].images.length > 0;

        s.addText(sl.title || `Slide ${i + 1}`, { x: 0.6, y: 0.5, w: 11.8, h: 0.9, fontFace: TITLE_FONT, fontSize: 26, color: INK });

        const bulletItems = (sl.bullets || []).map((b, bi) => ({ text: b, options: { bullet: { code: '2022' }, breakLine: bi < sl.bullets.length - 1, color: MUTED, paraSpaceAfter: 10 } }));
        const textW = hasImage ? 6.3 : 11.8;
        if (bulletItems.length) s.addText(bulletItems, { x: 0.6, y: 1.7, w: textW, h: 5.0, fontFace: BODY_FONT, fontSize: 15, valign: 'top' });

        if (hasImage) {
          const img = slides[i].images[0];
          const b64 = img.buffer.toString('base64');
          const mime = img.ext === 'jpg' || img.ext === 'jpeg' ? 'image/jpeg' : img.ext === 'gif' ? 'image/gif' : 'image/png';
          try {
            s.addImage({ data: `${mime};base64,${b64}`, x: 7.2, y: 1.7, w: 5.5, h: 4.5, sizing: { type: 'contain', w: 5.5, h: 4.5 } });
          } catch (e) { /* if an image is malformed, skip it rather than fail the whole deck */ }
        }

        s.addText(`${i + 1} / ${improved.slides.length}`, { x: 12.4, y: 7.1, w: 0.7, h: 0.3, fontFace: BODY_FONT, fontSize: 9, color: MUTED, align: 'right' });
      });

      const buf = await pres.write({ outputType: 'nodebuffer' });
      await auditLog(req.user.uid, req.user.username, 'ppt_polished', `${req.file.originalname} -> ${improved.slides.length} slides`, req.ip);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename="Polished_${req.file.originalname.replace(/\.pptx$/i, '')}.pptx"`);
      res.send(buf);
    } catch (e) {
      console.error('PPT polish error:', e.message); // full detail server-side only
      res.status(500).json({ error: 'Could not polish that file right now. Please try again.' });
    }
  });

  console.log('PPT Polish module: routes mounted');
};
