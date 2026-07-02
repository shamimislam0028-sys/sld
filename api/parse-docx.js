const JSZip = require('jszip');
const Busboy = require('busboy');

// ─────────────────── XML entity decoder ─────────────────────────
function decodeXml(s) {
  return s
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

// ─────────────────── Paragraph text extractor ───────────────────
function stripTags(s) {
  return s.replace(/<[^>]+>/g, '').trim();
}

function cleanOmmlPart(s) {
  return decodeXml(stripTags(s)).trim();
}

function ommlToLatex(xml) {
  let s = xml;
  const metaTags = [
    'm:rPr', 'm:sPr', 'm:fPr', 'm:radPr', 'm:dPr', 'm:naryPr',
    'm:accPr', 'm:groupChrPr', 'm:barPr', 'm:phantPr', 'w:rPr', 'w:pPr',
    'm:ctrlPr',
  ];
  for (const tag of metaTags) {
    s = s.replace(new RegExp(`<${tag}(?:\\s[^>]*)?>[\\s\\S]*?<\\/${tag}>`, 'g'), '');
    s = s.replace(new RegExp(`<${tag}(?:\\s[^>]*)?\\/>`, 'g'), '');
  }

  s = s.replace(/<(?:m:r|w:r)(?:\s[^>]*)?>([\s\S]*?)<\/(?:m:r|w:r)>/g, (_, inner) => {
    let text = '';
    const tRx = /<(?:m:t|w:t)(?:\s[^>]*)?>([\s\S]*?)<\/(?:m:t|w:t)>/g;
    let tm;
    while ((tm = tRx.exec(inner)) !== null) text += decodeXml(tm[1]);
    return text;
  });

  let prev = '';
  for (let iter = 0; iter < 60 && prev !== s; iter++) {
    prev = s;
    s = s.replace(
      /<m:f(?:\s[^>]*)?>(?:\s|<m:fPr[\s\S]*?<\/m:fPr>|<m:fPr[^>]*\/>)*<m:num(?:\s[^>]*)?>([\s\S]*?)<\/m:num>(?:\s|<[^>]+>)*<m:den(?:\s[^>]*)?>([\s\S]*?)<\/m:den>[\s\S]*?<\/m:f>/g,
      (_, num, den) => `\\frac{${ommlToLatex(num)}}{${ommlToLatex(den)}}`
    );
    s = s.replace(
      /<m:rad(?:\s[^>]*)?>(?:\s|<m:radPr[\s\S]*?<\/m:radPr>|<m:radPr[^>]*\/>)*(?:<m:deg(?:\s[^>]*)?>([\s\S]*?)<\/m:deg>|<m:deg(?:\s[^>]*)?\/>)(?:\s|<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>[\s\S]*?<\/m:rad>/g,
      (_, deg = '', body) => {
        const d = cleanOmmlPart(deg);
        return d ? `\\sqrt[${d}]{${ommlToLatex(body)}}` : `\\sqrt{${ommlToLatex(body)}}`;
      }
    );
    s = s.replace(
      /<m:sSup(?:\s[^>]*)?>(?:\s|<m:sSupPr[\s\S]*?<\/m:sSupPr>|<m:sSupPr[^>]*\/>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:\s|<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSup>/g,
      (_, base, exp) => `${ommlToLatex(base)}^{${ommlToLatex(exp)}}`
    );
    s = s.replace(
      /<m:sSub(?:\s[^>]*)?>(?:\s|<m:sSubPr[\s\S]*?<\/m:sSubPr>|<m:sSubPr[^>]*\/>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:\s|<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>[\s\S]*?<\/m:sSub>/g,
      (_, base, sub) => `${ommlToLatex(base)}_{${ommlToLatex(sub)}}`
    );
    s = s.replace(
      /<m:sSubSup(?:\s[^>]*)?>(?:\s|<m:sSubSupPr[\s\S]*?<\/m:sSubSupPr>|<m:sSubSupPr[^>]*\/>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:\s|<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>(?:\s|<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSubSup>/g,
      (_, base, sub, sup) => `${ommlToLatex(base)}_{${ommlToLatex(sub)}}^{${ommlToLatex(sup)}}`
    );
    s = s.replace(
      /<m:d(?:\s[^>]*)?>([\s\S]*?)<\/m:d>/g,
      (_, inner) => {
        const e = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/);
        return `(${ommlToLatex(e ? e[1] : inner)})`;
      }
    );
    s = s.replace(/<m:acc(?:\s[^>]*)?>([\s\S]*?)<\/m:acc>/g, (_, inner) => {
      const e = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/e>/);
      return ommlToLatex(e ? e[1] : inner);
    });
    for (const tag of ['m:e', 'm:num', 'm:den', 'm:sub', 'm:sup', 'm:deg', 'm:oMath', 'm:oMathPara', 'm:mr']) {
      s = s.replace(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'g'), '$1');
    }
  }
  return cleanOmmlPart(s);
}

function getParagraphLinesWithMath(pXml) {
  const lines = [];
  let currentLine = '';
  const xml = pXml.replace(/<w:tab(?:\s[^>]*)?\/?>/g, ' ');

  const tokenRx = /<w:br\b[^>]*\/?>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<m:oMath(?:\s[^>]*)?>([\s\S]*?)<\/m:oMath>|<m:oMathPara(?:\s[^>]*)?>([\s\S]*?)<\/m:oMathPara>/g;
  let m;
  while ((m = tokenRx.exec(xml)) !== null) {
    if (m[0].charAt(1) === 'w' && m[0].charAt(3) === 'b') {
      lines.push(currentLine.trim());
      currentLine = '';
    } else if (m[1] !== undefined) {
      currentLine += decodeXml(m[1] || '');
    } else if (m[2] !== undefined) {
      const latex = ommlToLatex(m[2] || '');
      if (latex) currentLine += ` $$${latex}$$ `;
    } else {
      const oMathRx = /<m:oMath(?:\s[^>]*)?>([\s\S]*?)<\/m:oMath>/g;
      let om;
      while ((om = oMathRx.exec(m[3] || '')) !== null) {
        const latex = ommlToLatex(om[1] || '');
        if (latex) currentLine += ` $$${latex}$$ `;
      }
    }
  }
  if (currentLine.trim()) lines.push(currentLine.trim());
  
  // স্পেস বাগ ফিক্স: ৩-৪টি অতিরিক্ত স্পেস থাকলে সেটিকে ১টি স্ট্যান্ডার্ড স্পেসে রূপান্তর করবে
  return lines.map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

// ─────────────────── Smart Engine For All Docx ───────────────────
const LETTER_MAP = { 'ক': 'A', '১': 'A', 'খ': 'B', '২': 'B', 'গ': 'C', '৩': 'C', 'ঘ': 'D', '৪': 'D' };
function normLetter(ch) { return LETTER_MAP[ch] || ch.toUpperCase(); }

const OPT = {
  A: /^[(\[]?\s*[Aaক১]\s*[.)\]\u0964\-।]\s*/,
  B: /^[(\[]?\s*[Bbখ২]\s*[.)\]\u0964\-।]\s*/,
  C: /^[(\[]?\s*[Ccগ৩]\s*[.)\]\u0964\-]\s*/,
  D: /^[(\[]?\s*[Ddঘ৪]\s*[.)\]\u0964\-]\s*/,
};
const INLINE_OPT = /[(\[]?([A-Da-dক-ঘ১-৪])\s*[.)\]\u0964\-।]\s+/g;
const ANS = /^(?:Correct\s*Option|Correct|Answer|Ans|উত্তর|সঠিক\s*উত্তর|সঠিক)\s*[:.\-।\s]*([A-Da-dক-ঘ১-৪])/i;
const NUMQ = /^(?:Question\s*[\d০-৯]+|[Qq](?:\.|uestion)?\s*[\d০-৯]+|[\d০-৯]+\s*[.)\]\u0964।\-]\s*|\([\d০-৯]+\)|\[[\d০-৯]+\])/;
const ROMAN = /^(?:\s*|[(ঁঁ]*)([ivxlcdm]+|[i০-৯]+)\s*[.)\]\u0964\-।]/i;
const PASSAGE = /(উদ্দীপক|অনুচ্ছেদ|দৃশ্যকল্প|উদ্ধৃত(?:াংশ)?|কবিতাংশ|চিত্র|চিত্রে|নিচের\s*|নিচের\s*.*(?:পড়|লক্ষ|পড়ো)|প্রশ্নের?\s*উত্তর\s*দাও)/;

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
  let activePassage = ''; 

  const flushCurrent = () => {
    if (cur && cur.question && (cur.optionA || cur.optionB || cur.optionC || cur.optionD)) {
      if (!cur.passage && activePassage) {
        cur.passage = activePassage;
      }
      mcqs.push(cur);
    }
    cur = null;
  };

  for (let i = 0; i < lineData.length; i++) {
    const obj = lineData[i];
    const line = (obj.text || '').replace(/\s+/g, ' ').trim();
    if (!line) continue;

    // ১) Answer লাইন চেক
    const corr = getCorrect(line);
    if (corr) {
      if (cur) {
        cur.correct = corr;
        flushCurrent();
      }
      continue;
    }

    // ২) Option লাইন চেক (একক বা ইনলাইন অপশন)
    const inlineOpts = parseInlineOptions(line);
    const singleOptLetter = getOptionLetter(line);

    if (inlineOpts || singleOptLetter) {
      if (!cur) {
        cur = { passage: activePassage, question: 'প্রশ্ন পাওয়া যায়নি', optionA: '', optionB: '', optionC: '', optionD: '', correct: 'A' };
      }
      if (inlineOpts) {
        if (inlineOpts.A) cur.optionA = inlineOpts.A;
        if (inlineOpts.B) cur.optionB = inlineOpts.B;
        if (inlineOpts.C) cur.optionC = inlineOpts.C;
        if (inlineOpts.D) cur.optionD = inlineOpts.D;
      } else {
        cur['option' + singleOptLetter] = stripOpt(line, singleOptLetter);
      }
      continue;
    }

    // ৩) উদ্দীপক (Passage) চেক
    const isRealPassage = PASSAGE.test(line) && !line.endsWith('?') && !line.includes('কোনটি সঠিক') && !line.includes('কোনটি') && !obj.isNumbered;
    if (isRealPassage) {
      flushCurrent();
      activePassage = line;
      
      while (i + 1 < lineData.length) {
        const nextLine = (lineData[i + 1].text || '').replace(/\s+/g, ' ').trim();
        if (!nextLine || NUMQ.test(nextLine) || lineData[i + 1].isNumbered || getOptionLetter(nextLine) || parseInlineOptions(nextLine) || ANS.test(nextLine)) {
          break;
        }
        activePassage += '\n' + nextLine; // উদ্দীপকেও লাইন ব্রেক বজায় থাকবে
        i++;
      }
      continue;
    }

    // ৪) নতুন প্রশ্ন (Numbered Question) শুরুর ডিটেকশন
    const isContinuation = line.includes('কোনটি সঠিক') || line.includes('নিচের কোনটি') || ROMAN.test(line);
    const isNewQuestionStart = (NUMQ.test(line) || obj.isNumbered) && !isContinuation;

    if (isNewQuestionStart) {
      flushCurrent();

      if (!line.includes('উদ্দীপক') && !line.includes('চিত্র') && !line.includes('নং')) {
        activePassage = '';
      }

      let cleanedQuestion = NUMQ.test(line) ? line.replace(NUMQ, '').replace(/^[:।.\-\s]+/, '').trim() : line;
      if (!cleanedQuestion) cleanedQuestion = line;

      cur = { passage: activePassage, question: cleanedQuestion, optionA: '', optionB: '', optionC: '', optionD: '', correct: 'A' };
      continue;
    }

    // ৫) প্রশ্নের ধারাবাহিক অংশ (রোমান সংখ্যা এবং 'নিচের কোনটি সঠিক?' টেক্সটকে নিখুঁতভাবে নিচে নিচে সাজাবে)
    if (cur) {
      cur.question = cur.question ? cur.question + '\n' + line : line; // এখানে \n দেওয়া হয়েছে যাতে ফ্রন্টএন্ডে নিচে নিচে দেখায়
    } else {
      if (!isContinuation) {
        activePassage = ''; 
      }
      cur = { passage: activePassage, question: line, optionA: '', optionB: '', optionC: '', optionD: '', correct: 'A' };
    }
  }

  flushCurrent();
  return mcqs.filter(m => m.question && (m.optionA || m.optionB || m.optionC || m.optionD));
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
      getParagraphLinesWithMath(pXml).forEach(line => lineData.push({ text: line, isNumbered }));
    }

    if (!lineData.length) return res.status(422).json({ error: 'No text found in DOCX.' });

    const mcqs = parseParagraphsToMcqs(lineData);
    if (!mcqs.length) return res.status(422).json({ error: 'No MCQs found.' });

    return res.status(200).json({ mcqs });
  } catch (err) {
    console.error('parse-docx error:', err);
    return res.status(500).json({ error: err.message || 'DOCX parsing failed.' });
  }
};
