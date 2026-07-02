/**
 * server.js — Express API for the MCQ PowerPoint Generator.
 *
 * Endpoints:
 *   POST /api/generate   multipart/form-data
 *     fields:
 *       - mcqs        JSON string: [{question,optionA,optionB,optionC,optionD,correct}]
 *       - mnemonic    string (course mnemonic; stored in metadata)
 *       - icon        file (PNG/JPG/SVG) — replaces the master background image
 *     returns: the generated .pptx as a download
 *
 *   GET  /api/health    liveness probe
 *
 * The heavy lifting lives in src/pptEngine.js, which reuses the uploaded
 * master template VERBATIM. To swap templates later, replace
 * templates/master.pptx (and, if shape ids change, the TEMPLATE map in
 * pptEngine.js) — server logic stays the same.
 */

const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const Jimp = require('jimp');
const { generatePptx } = require('./src/pptEngine');

const app = express();
const PORT = process.env.PORT || 4000;

const TEMPLATE_PATH = path.join(__dirname, 'templates', 'master.pptx');

app.use(cors());
app.use(express.json({ limit: '5mb' }));

// in-memory upload (icons are small)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 8 * 1024 * 1024 }, // 8 MB
});

/* ----------------------------- helpers ---------------------------- */

function validateMcqs(mcqs) {
  if (!Array.isArray(mcqs) || mcqs.length === 0) {
    return 'At least one question is required.';
  }
  for (let i = 0; i < mcqs.length; i++) {
    const m = mcqs[i];
    const n = i + 1;
    if (!m.question || !m.question.trim()) return `Question ${n}: question text is empty.`;
    if (!m.optionA || !m.optionA.trim()) return `Question ${n}: option A is empty.`;
    if (!m.optionB || !m.optionB.trim()) return `Question ${n}: option B is empty.`;
    if (!m.optionC || !m.optionC.trim()) return `Question ${n}: option C is empty.`;
    if (!m.optionD || !m.optionD.trim()) return `Question ${n}: option D is empty.`;
    if (!['A', 'B', 'C', 'D'].includes(m.correct))
      return `Question ${n}: correct answer must be A, B, C, or D.`;
  }
  return null;
}

/**
 * Normalize an uploaded image to PNG (the layout boxes are filled with PNGs).
 * SVG and JPG are rasterized/transcoded via sharp so the media part stays a
 * valid PNG and renders identically across PowerPoint versions.
 */
async function imageToPng(file) {
  if (!file) return null;
  try {
    const image = await Jimp.read(file.buffer);
    return await image.getBufferAsync(Jimp.MIME_PNG);
  } catch (err) {
    throw new Error('Could not process the image: ' + err.message);
  }
}

/* ------------------------------ routes ---------------------------- */

app.get(['/api/health', '/health'], (_req, res) => res.json({ ok: true }));


/* ========== DOCX → MCQ JSON (JavaScript port of the Python script) ========== */
const JSZip = require('jszip');

const docxUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 4 * 1024 * 1024 }, // 4 MB — Vercel free plan limit
});

/**
 * Convert OMML (Office Math Markup Language) XML → LaTeX string.
 * Handles: fractions, sqrt, superscript, subscript, delimiters, accents.
 * Uses an inside-out iterative approach: processes innermost elements first.
 */
function ommlToLatex(xml) {
  let s = xml;

  // 1. Strip property/metadata elements (no content value)
  const metaTags = ['m:rPr','m:sPr','m:fPr','m:radPr','m:dPr','m:naryPr',
                     'm:accPr','m:groupChrPr','m:barPr','m:phantPr','w:rPr','w:pPr'];
  for (const t of metaTags) {
    s = s.replace(new RegExp(`<${t}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${t}>`, 'g'), '');
    s = s.replace(new RegExp(`<${t}(?:\\s[^>]*)?\\/>`, 'g'), '');
  }

  // 2. Collapse m:r / w:r runs to just their text
  s = s.replace(/<(?:m:r|w:r)(?:\s[^>]*)?>([\s\S]*?)<\/(?:m:r|w:r)>/g, (_, inner) => {
    const tm = inner.match(/<(?:m:t|w:t)(?:\s[^>]*)?>([^<]*)<\/(?:m:t|w:t)>/);
    return tm ? tm[1] : '';
  });

  // 3. Iteratively convert OMML structures → LaTeX (innermost first)
  let prev = '';
  for (let iter = 0; iter < 40 && prev !== s; iter++) {
    prev = s;

    // Fraction  \frac{num}{den}
    s = s.replace(
      /<m:f(?:\s[^>]*)?>(?:<[^>]+>)*<m:num(?:\s[^>]*)?>([\s\S]*?)<\/m:num>(?:<[^>]+>)*<m:den(?:\s[^>]*)?>([\s\S]*?)<\/m:den>[\s\S]*?<\/m:f>/,
      (_, n, d) => `\\frac{${clean(n)}}{${clean(d)}}`
    );

    // Radical / sqrt
    s = s.replace(
      /<m:rad(?:\s[^>]*)?>(?:<[^>]+>)*<m:deg(?:\s[^>]*)?>([\s\S]*?)<\/m:deg>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>[\s\S]*?<\/m:rad>/,
      (_, deg, rad) => {
        const d = clean(deg).trim();
        return d ? `\\sqrt[${d}]{${clean(rad)}}` : `\\sqrt{${clean(rad)}}`;
      }
    );

    // Superscript  base^{exp}
    s = s.replace(
      /<m:sSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSup>/,
      (_, base, exp) => `${clean(base)}^{${clean(exp)}}`
    );

    // Subscript  base_{sub}
    s = s.replace(
      /<m:sSub(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>[\s\S]*?<\/m:sSub>/,
      (_, base, sub) => `${clean(base)}_{${clean(sub)}}`
    );

    // Sub+Superscript
    s = s.replace(
      /<m:sSubSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSubSup>/,
      (_, base, sub, sup) => `${clean(base)}_{${clean(sub)}}^{${clean(sup)}}`
    );

    // Delimiter (parentheses / brackets)
    s = s.replace(
      /<m:d(?:\s[^>]*)?>([\s\S]*?)<\/m:d>/,
      (_, inner) => {
        const eMatch = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/);
        return eMatch ? `(${clean(eMatch[1])})` : `(${clean(inner)})`;
      }
    );

    // Accent / overline (just keep inner)
    s = s.replace(/<m:acc(?:\s[^>]*)?>([\s\S]*?)<\/m:acc>/,
      (_, inner) => { const e = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/); return e ? clean(e[1]) : clean(inner); }
    );

    // Structural wrappers — strip tag, keep content
    for (const tag of ['m:e','m:num','m:den','m:sub','m:sup','m:deg','m:oMath','m:oMathPara','m:mr']) {
      s = s.replace(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'g'), '$1');
    }

    // Remove any remaining unknown m: tags (keep their text content via stripTags later)
    s = s.replace(/<m:[a-zA-Z]+(?:\s[^>]*)?>[\s\S]*?<\/m:[a-zA-Z]+>/g, (m) =>
      m.replace(/<[^>]+>/g, '')
    );
  }

  return s.replace(/<[^>]+>/g, '').trim();
}

/** Remove all XML tags, leaving only text. */
function clean(s) {
  return s.replace(/<[^>]+>/g, '').trim();
}

/**
 * Extract all text from a <w:p> paragraph XML fragment,
 * and split into individual lines where <w:br/> occurs.
 */
function getParagraphLines(pXml) {
  let currentLine = '';
  const lines = [];

  // First, handle any w:tabs by replacing them with spaces
  let xml = pXml.replace(/<w:tab(?:\s[^>]*)?\/>/g, ' ');
  
  const childRx = /<(w:r|m:oMath|m:oMathPara)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/g;
  let m;

  while ((m = childRx.exec(xml)) !== null) {
    const tag     = m[1];
    const content = m[2];

    if (tag === 'w:r') {
      // Check for <w:br/> → split line
      const brMatches = content.match(/<w:br\b/g);
      const tRx = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let tm, textPart = '';
      while ((tm = tRx.exec(content)) !== null) textPart += tm[1];
      
      if (brMatches) {
        // There are <w:br/> tags
        if (currentLine.trim() || textPart.trim()) {
          lines.push((currentLine + textPart).trim());
        }
        for (let i = 0; i < brMatches.length - 1; i++) {
          lines.push('');
        }
        currentLine = '';
      } else {
        currentLine += textPart;
      }
    } else if (tag === 'm:oMath') {
      const latex = ommlToLatex(content);
      if (latex) currentLine += ` $$${latex}$$ `;
    } else if (tag === 'm:oMathPara') {
      const oMathRx = /<m:oMath(?:\s[^>]*)?>([\s\S]*?)<\/m:oMath>/g;
      let om;
      while ((om = oMathRx.exec(content)) !== null) {
        const latex = ommlToLatex(om[1]);
        if (latex) currentLine += ` $$${latex}$$ `;
      }
    }
  }

  if (currentLine.trim()) {
    lines.push(currentLine.trim());
  }

  // Normalize whitespace in each line
  return lines.map(line => line.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

function getParagraphText(pXml) {
  // Legacy function, keep for compatibility
  return getParagraphLines(pXml).join('\n');
}

/**
 * Port of Python docx_to_mcq_json() MCQ-detection logic.
 * Detects:
 *   Questions : line starts with digit+separator, OR ends with "?"
 *   Options   : A. A) B. B) C. C) D. D) (case-insensitive)
/** True if a block of text looks like an উদ্দীপক/passage (not a stray heading or metadata). */
function looksLikePassage(text) {
  const cleanText = text.trim();
  if (!cleanText) return false;

  // 1. Check strong passage keywords
  const rxPassageKeywords = /(উদ্দীপক|অনুচ্ছেদ|দৃশ্যকল্প|উদ্ধৃত|কবিতাংশ|পড়ে|পড়ো|পড়ুন|পাঠ|উত্তর\s*দাও)/;
  if (rxPassageKeywords.test(cleanText)) return true;

  // 2. If it is a multi-line text or a long paragraph, and NOT a metadata header line
  const rxMetadata = /^(শ্রেণি|বিষয়|শ্রেণী|বিষয়|অধ্যায়|অধ্যায়|সময়|সময়|পূর্ণমান|মান|কোড|পরীক্ষা|তারিখ)\s*[:।]/i;
  
  if (cleanText.length > 50 && !rxMetadata.test(cleanText)) {
    return true;
  }

  // 3. Fallback checking for multiple lines (e.g. poetry or multi-paragraph passage)
  const lines = cleanText.split('\n').filter(Boolean);
  return lines.length >= 2 && cleanText.length > 60;
}

/**
 * Two-pass MCQ parser.
 *
 * Pass 1 – Segment every line into either a "passage" segment or a
 *           "question" segment (question line + its options/answer).
 * Pass 2 – For each question segment, look up the most recent real
 *           passage segment and prepend it to the question text.
 *
 * This avoids the single-pass bug where passage accumulation fails if
 * any earlier line accidentally triggers question detection.
 */
const LETTER_MAP = { 'ক': 'A', '১': 'A', 'খ': 'B', '২': 'B', 'গ': 'C', '৩': 'C', 'ঘ': 'D', '৪': 'D' };
function normLetter(ch) { return LETTER_MAP[ch] || ch.toUpperCase(); }

const OPT = {
  A: /^[(\[]?\s*[Aaক১]\s*[.)\]\u0964\-]\s+/,
  B: /^[(\[]?\s*[Bbখ২]\s*[.)\]\u0964\-]\s+/,
  C: /^[(\[]?\s*[Ccগ৩]\s*[.)\]\u0964\-]\s+/,
  D: /^[(\[]?\s*[Ddঘ৪]\s*[.)\]\u0964\-]\s+/,
};
const INLINE_OPT = /[(\[]?([A-Da-dক-ঘ১-৪])\s*[.)\]\u0964\-]\s+/g;
const ANS = /^(?:Correct\s*Option|Correct|Answer|Ans|উত্তর|সঠিক\s*উত্তর|সঠিক)\s*[:.\-।]?\s*([A-Da-dক-ঘ১-৪])/i;
const NUMQ = /^(?:Question\s*[\d০-৯]+|[Qq](?:\.|uestion)?\s*[\d০-৯]+|[\d০-৯]+\s*[.)\]\u0964]|\([\d০-৯]+\)|\[[\d০-৯]+\])/;
const ROMAN = /^\s*(?:[ivxlcdm]+)\s*[.)]\s+/i;
const PASSAGE = /(উদ্দীপক|অনুচ্ছেদ|দৃশ্যকল্প|উদ্ধৃত(?:াংশ)?|কবিতাংশ|চিত্র|নিচের\s*.*(?:পড়|লক্ষ)|প্রশ্নের?\s*উত্তর\s*দাও)/;

function getOptionLetter(line) {
  for (const L of ['A', 'B', 'C', 'D']) if (OPT[L].test(line)) return L;
  return null;
}
function stripOpt(line, L) { return line.replace(OPT[L], '').trim(); }

function parseInlineOptions(line) {
  INLINE_OPT.lastIndex = 0;
  const matches = [];
  let m;
  while ((m = INLINE_OPT.exec(line)) !== null) {
    matches.push({ full: m[0], letter: normLetter(m[1]), index: m.index });
  }
  if (matches.length < 2) return null;
  const parsed = {};
  for (let i = 0; i < matches.length; i++) {
    const cur = matches[i], next = matches[i + 1];
    const start = cur.index + cur.full.length;
    const end = next ? next.index : line.length;
    const val = line.slice(start, end).trim();
    if ('ABCD'.includes(cur.letter) && val) parsed[cur.letter] = val;
  }
  return Object.keys(parsed).length >= 2 ? parsed : null;
}
function getCorrect(line) {
  const m = line.match(ANS);
  return m ? normLetter(m[1]) : null;
}

// Format-agnostic MCQ parser (same logic as api/parse-docx.js).
function parseParagraphsToMcqs(lineData) {
  const mcqs = [];
  let cur = null;
  let pending = '';
  let inPassage = false;

  const blank = (q) => ({ passage: pending, question: q || '', optionA: '', optionB: '', optionC: '', optionD: '', correct: 'A' });
  const hasOpts = (m) => !!(m && (m.optionA || m.optionB || m.optionC || m.optionD));
  const flush = () => { if (cur && cur.question && hasOpts(cur)) mcqs.push(cur); };

  for (const obj of lineData) {
    const line = (obj.text || '').replace(/\s+/g, ' ').trim();
    if (!line) continue;

    const corr = getCorrect(line);
    if (corr && cur) { cur.correct = corr; flush(); cur = null; pending = ''; inPassage = false; continue; }

    const numberedByStyle = !!obj.isNumbered && !getOptionLetter(line);
    const numberedByText = NUMQ.test(line) && !ROMAN.test(line);
    if (numberedByStyle || numberedByText) {
      let q = numberedByText ? line.replace(NUMQ, '').replace(/^[:।.\-\s]+/, '').trim() : line;
      if (!q) q = line;
      flush();
      cur = blank(q);
      pending = ''; inPassage = false;
      continue;
    }

    const inl = cur ? parseInlineOptions(line) : null;
    if (inl) {
      if (inl.A) cur.optionA = inl.A;
      if (inl.B) cur.optionB = inl.B;
      if (inl.C) cur.optionC = inl.C;
      if (inl.D) cur.optionD = inl.D;
      inPassage = false; continue;
    }

    const optL = cur ? getOptionLetter(line) : null;
    if (optL) { cur['option' + optL] = stripOpt(line, optL); inPassage = false; continue; }

    if (PASSAGE.test(line) && !line.endsWith('?') && !obj.isNumbered && !hasOpts(cur)) {
      pending = pending ? pending + ' ' + line : line;
      inPassage = true; continue;
    }
    if (inPassage && !obj.isNumbered) {
      pending = pending ? pending + ' ' + line : line;
      continue;
    }

    if (cur && !hasOpts(cur)) {
      cur.question = cur.question ? cur.question + ' ' + line : line;
      continue;
    }

    flush();
    cur = blank(line);
    pending = '';
  }

  flush();
  return mcqs.filter(m => m.question && hasOpts(m));
}


/* ─────────────────── POST /api/parse-docx ─────────────────── */
app.post(
  ['/api/parse-docx', '/parse-docx'],
  docxUpload.single('docx'),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'No DOCX file uploaded.' });
      }

      // Unzip the DOCX (it is a ZIP archive)
      const zip = await JSZip.loadAsync(req.file.buffer);
      const docEntry = zip.file('word/document.xml');
      if (!docEntry) {
        return res.status(422).json({ error: 'Invalid DOCX: word/document.xml not found.' });
      }

      const docXml = await docEntry.async('string');

      // Extract all lines in document order, and track if they are from a numbered paragraph
      const lineData = []; // Each item { text: string, isNumbered: boolean }
      const pRx = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
      let pm;
      while ((pm = pRx.exec(docXml)) !== null) {
        const pXml = pm[1];
        const isNumbered = /<w:numPr>/.test(pXml);
        const lines = getParagraphLines(pXml);
        lines.forEach(line => {
          lineData.push({ text: line, isNumbered });
        });
      }

      console.log('DEBUG: Extracted line data:', lineData);

      if (!lineData.length) {
        return res.status(422).json({ error: 'No text found in DOCX.' });
      }

      const mcqs = parseParagraphsToMcqs(lineData);
      
      console.log('DEBUG: Parsed MCQs:', mcqs);

      if (!mcqs.length) {
        return res.status(422).json({
          error:
            'No MCQs found. Make sure each question is followed by options A. B. C. D. and a correct answer.',
        });
      }

      return res.json({ mcqs });
    } catch (err) {
      console.error(err);
      return res.status(500).json({ error: err.message || 'DOCX parsing failed.' });
    }
  }
);


app.post(
  ['/api/generate', '/generate'],
  upload.fields([
    { name: 'icon', maxCount: 1 },
    { name: 'mnemonic', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      // Parse + validate inputs
      let mcqs;
      try {
        mcqs = JSON.parse(req.body.mcqs || '[]');
      } catch {
        return res.status(400).json({ error: 'mcqs must be valid JSON.' });
      }

      const iconFile = req.files?.icon?.[0] || null;
      const mnemonicFile = req.files?.mnemonic?.[0] || null;

      if (!mnemonicFile) {
        return res.status(400).json({ error: 'Course mnemonic image is required.' });
      }
      if (!iconFile) {
        return res.status(400).json({ error: 'Course icon is required.' });
      }

      const vErr = validateMcqs(mcqs);
      if (vErr) return res.status(400).json({ error: vErr });

      if (!fs.existsSync(TEMPLATE_PATH)) {
        return res
          .status(500)
          .json({ error: 'Server template missing: templates/master.pptx' });
      }

      const template = fs.readFileSync(TEMPLATE_PATH);
      const iconPng = await imageToPng(iconFile);
      const mnemonicPng = await imageToPng(mnemonicFile);

      const buffer = await generatePptx(template, mcqs, {
        iconBuffer: iconPng,
        mnemonicBuffer: mnemonicPng,
      });

      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      );
      res.setHeader(
        'Content-Disposition',
        'attachment; filename="MCQ_deck.pptx"'
      );
      return res.send(buffer);
    } catch (err) {
      console.error(err);
      return res.status(500).json({ error: err.message || 'Generation failed.' });
    }
  }
);

if (require.main === module) {
  // Serve the built frontend in production (optional) only when running locally
  const FRONTEND_DIST = path.join(__dirname, '..', 'frontend', 'dist');
  if (fs.existsSync(FRONTEND_DIST)) {
    app.use(express.static(FRONTEND_DIST));
    app.get('*', (_req, res) =>
      res.sendFile(path.join(FRONTEND_DIST, 'index.html'))
    );
  }

  app.listen(PORT, () => {
    console.log(`MCQ PPTX generator API running on http://localhost:${PORT}`);
  });
}

module.exports = {
  app,
  parseParagraphsToMcqs,
  getParagraphText,
  getParagraphLines
};
