const JSZip = require('jszip');
const Busboy = require('busboy');
const { extractMcqsWithGemini, lineDataToText, hasGeminiKey } = require('./gemini-extract');

// ─────────────────── XML entity decoder ─────────────────────────
function decodeXml(s) {
  return s
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

// ─────────────────── Paragraph text extractor ───────────────────
// Captures BOTH normal text (<w:t>) and math text (<m:t>) in document
// order, splitting on soft line breaks (<w:br>). Light OMML hints make
// fractions / roots / powers human-readable (3/5, √10, p^2) instead of
// silently dropping math-only questions and options.
function getParagraphLines(pXml) {
  const lines = [];
  let currentLine = '';
  let xml = pXml.replace(/<w:tab(?:\s[^>]*)?\/?>/g, ' ');

  // OMML → readable text hints
  xml = xml
    .replace(/<\/m:num>/g, '</m:num><m:t>/</m:t>')              // fraction: num / den
    .replace(/<m:rad(\s[^>]*)?>/g, '<m:rad$1><m:t>\u221A</m:t>') // radical: √
    .replace(/<m:sup>/g, '<m:t>^</m:t><m:sup>')                 // superscript / power
    .replace(/<m:sub>/g, '<m:t>_</m:t><m:sub>');                // subscript

  const tokenRx = /<w:br\b[^>]*\/?>|<(w:t|m:t)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/g;
  let m;
  while ((m = tokenRx.exec(xml)) !== null) {
    if (m[0].charAt(1) === 'w' && m[0].charAt(3) === 'b') { // <w:br>
      lines.push(currentLine.trim());
      currentLine = '';
    } else {
      currentLine += decodeXml(m[2] || '');
    }
  }
  if (currentLine.trim()) lines.push(currentLine.trim());
  return lines.map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

// ─────────────────── MCQ parser (format-agnostic) ───────────────
// Handles, in any combination:
//   • question numbering: Word auto-number (numPr) OR typed "১।" "1." "1)" "(1)" "[1]" "Q1" "Question 1"
//   • options: Latin A–D, Bengali ক/খ/গ/ঘ, Bengali digits ১–৪ — one per line OR several inline
//   • answers: "উত্তর", "সঠিক উত্তর", "Answer", "Ans", "Correct" → letter in any of the above scripts
//   • passages / উদ্দীপক stimulus blocks shared by following questions
//   • roman sub-items i. ii. iii. kept inside the question stem
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

function parseParagraphsToMcqs(lineData) {
  const mcqs = [];
  let cur = null;
  let pending = '';      // accumulated passage / উদ্দীপক text
  let inPassage = false;

  const blank = (q) => ({ passage: pending, question: q || '', optionA: '', optionB: '', optionC: '', optionD: '', correct: 'A' });
  const hasOpts = (m) => !!(m && (m.optionA || m.optionB || m.optionC || m.optionD));
  const flush = () => { if (cur && cur.question && hasOpts(cur)) mcqs.push(cur); };

  for (const obj of lineData) {
    const line = (obj.text || '').replace(/\s+/g, ' ').trim();
    if (!line) continue;

    // 1) ANSWER line → close the current question
    const corr = getCorrect(line);
    if (corr && cur) { cur.correct = corr; flush(); cur = null; pending = ''; inPassage = false; continue; }

    // 2) NUMBERED question start (Word auto-number OR typed number; roman i./ii. excluded)
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

    // 3) INLINE options (several markers on one line)
    const inl = cur ? parseInlineOptions(line) : null;
    if (inl) {
      if (inl.A) cur.optionA = inl.A;
      if (inl.B) cur.optionB = inl.B;
      if (inl.C) cur.optionC = inl.C;
      if (inl.D) cur.optionD = inl.D;
      inPassage = false; continue;
    }

    // 4) SINGLE option per line
    const optL = cur ? getOptionLetter(line) : null;
    if (optL) { cur['option' + optL] = stripOpt(line, optL); inPassage = false; continue; }

    // 5) PASSAGE / stimulus intro (not a question, not numbered, before options)
    if (PASSAGE.test(line) && !line.endsWith('?') && !obj.isNumbered && !hasOpts(cur)) {
      pending = pending ? pending + ' ' + line : line;
      inPassage = true; continue;
    }
    if (inPassage && !obj.isNumbered) {
      pending = pending ? pending + ' ' + line : line;
      continue;
    }

    // 6) continuation of the current question stem (roman items, "নিচের কোনটি সঠিক?", wrapped text)
    if (cur && !hasOpts(cur)) {
      cur.question = cur.question ? cur.question + ' ' + line : line;
      continue;
    }

    // 7) fallback — an orphan line starts a new (un-numbered) question
    flush();
    cur = blank(line);
    pending = '';
  }

  flush();
  return mcqs.filter(m => m.question && hasOpts(m));
}

// ─────────────────── Vercel serverless handler ──────────────────
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Content-Type', 'application/json');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed.' });

  try {
    const fileBuffer = await new Promise((resolve, reject) => {
      const contentType = req.headers['content-type'] || '';
      if (!contentType.includes('multipart/form-data')) {
        return reject(new Error('Expected multipart/form-data'));
      }
      const bb = Busboy({ headers: req.headers, limits: { fileSize: 4 * 1024 * 1024 } });
      let fileData = null;
      bb.on('file', (_name, stream) => {
        const chunks = [];
        stream.on('data', c => chunks.push(c));
        stream.on('end', () => { fileData = Buffer.concat(chunks); });
      });
      bb.on('finish', () => {
        if (!fileData) return reject(new Error('No DOCX file uploaded.'));
        resolve(fileData);
      });
      bb.on('error', reject);
      req.pipe(bb);
    });

    const zip = await JSZip.loadAsync(fileBuffer);
    const docEntry = zip.file('word/document.xml');
    if (!docEntry) return res.status(422).json({ error: 'Invalid DOCX: word/document.xml not found.' });

    const docXml = await docEntry.async('string');
    const lineData = [];
    const pRx = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
    let pm;
    while ((pm = pRx.exec(docXml)) !== null) {
      const pXml = pm[0].replace(/^<w:p[^>]*>/, '').replace(/<\/w:p>$/, '');
      const isNumbered = /<w:numPr>/.test(pXml);
      getParagraphLines(pXml).forEach(line => lineData.push({ text: line, isNumbered }));
    }

    if (!lineData.length) return res.status(422).json({ error: 'No text found in DOCX.' });

    // Primary: extract with Gemini (handles any format). Fallback: regex parser,
    // so the app keeps working if the key is missing or the free quota is exhausted.
    let mcqs = [];
    let source = 'gemini';
    if (hasGeminiKey()) {
      try {
        mcqs = await extractMcqsWithGemini(lineDataToText(lineData));
      } catch (aiErr) {
        console.error('Gemini extraction failed, falling back to regex:', aiErr.message);
        mcqs = parseParagraphsToMcqs(lineData);
        source = 'regex-fallback';
      }
    } else {
      mcqs = parseParagraphsToMcqs(lineData);
      source = 'regex-no-key';
    }

    if (!mcqs.length) return res.status(422).json({ error: 'No MCQs found. Make sure each question is followed by options A. B. C. D. and a correct answer.' });

    return res.status(200).json({ mcqs, source });
  } catch (err) {
    console.error('parse-docx error:', err);
    return res.status(500).json({ error: err.message || 'DOCX parsing failed.' });
  }
};

// Exported for local reuse / testing
module.exports.getParagraphLines = getParagraphLines;
module.exports.parseParagraphsToMcqs = parseParagraphsToMcqs;
