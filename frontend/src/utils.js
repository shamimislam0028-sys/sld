// utils.js — shared, framework-agnostic helpers
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

export const LETTERS = ['A', 'B', 'C', 'D'];
export const STORAGE_KEY = 'mcq_pptx_state_v1';

export function blankMcq() {
  return {
    id: crypto.randomUUID(),
    passage: '',
    question: '',
    optionA: '',
    optionB: '',
    optionC: '',
    optionD: '',
    correct: 'A',
  };
}

/** Validate a single MCQ; returns array of human-readable error strings. */
export function validateMcq(m, index) {
  const errs = [];
  const n = index + 1;
  if (!m.question?.trim()) errs.push(`Q${n}: question is empty`);
  if (!m.optionA?.trim()) errs.push(`Q${n}: option A is empty`);
  if (!m.optionB?.trim()) errs.push(`Q${n}: option B is empty`);
  if (!m.optionC?.trim()) errs.push(`Q${n}: option C is empty`);
  if (!m.optionD?.trim()) errs.push(`Q${n}: option D is empty`);
  if (!LETTERS.includes(m.correct)) errs.push(`Q${n}: pick a correct answer`);
  return errs;
}

/** Validate everything before generation. */
export function validateAll({ mnemonicFile, iconFile, mcqs }) {
  const errs = [];
  if (!mnemonicFile) errs.push('Course mnemonic image is required');
  if (!iconFile) errs.push('Course icon is required');
  if (!mcqs?.length) errs.push('Add at least one question');
  mcqs?.forEach((m, i) => errs.push(...validateMcq(m, i)));
  return errs;
}

/** Persist UI state to localStorage (files can't serialize, only questions). */
export function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ mcqs: state.mcqs }));
  } catch {
    /* quota / private mode — ignore */
  }
}

export function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    // re-issue ids in case of old saves
    if (Array.isArray(parsed.mcqs)) {
      parsed.mcqs = parsed.mcqs.map((m) => ({ ...blankMcq(), ...m, id: m.id || crypto.randomUUID() }));
    }
    return parsed;
  } catch {
    return null;
  }
}

export function clearState() {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    /* ignore */
  }
}

/** Export MCQs to a downloadable JSON file. */
export function exportJson(mcqs) {
  const clean = mcqs.map(({ passage, question, optionA, optionB, optionC, optionD, correct }) => ({
    passage: passage ?? '',
    question,
    optionA,
    optionB,
    optionC,
    optionD,
    correct,
  }));
  const blob = new Blob([JSON.stringify(clean, null, 2)], {
    type: 'application/json',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'mcq-questions.json';
  a.click();
  URL.revokeObjectURL(url);
}

/** Parse imported JSON into normalized MCQs. Accepts array or {mcqs:[...]}. */
export function parseImportedJson(text) {
  const data = JSON.parse(text);
  const arr = Array.isArray(data) ? data : data.mcqs;
  if (!Array.isArray(arr)) throw new Error('JSON must be an array of questions.');
  return arr.map((m) => ({
    ...blankMcq(),
    passage: m.passage ?? '',
    question: m.question ?? '',
    optionA: m.optionA ?? m.a ?? '',
    optionB: m.optionB ?? m.b ?? '',
    optionC: m.optionC ?? m.c ?? '',
    optionD: m.optionD ?? m.d ?? '',
    correct: LETTERS.includes(m.correct) ? m.correct : 'A',
  }));
}

/** Check if a text item qualifies as a fraction numerator or denominator */
function isFractionPart(text, width) {
  const clean = text.trim();
  if (!clean) return false;
  if (width > 150) return false;
  
  // Must NOT contain general Bengali non-digit characters
  if (/[\u0980-\u09e5\u09f0-\u09ff]/.test(clean)) {
    return false;
  }
  
  // Must consist only of math/numeric symbols
  return /^[0-9a-zA-Z\+\-\*\/\^\(\)\{\}\s\u09e6-\u09ef\±\√\θ\β\α\=,\.\!\?\[\]\\\|]+$/.test(clean);
}

/** Reconstruct stacked fractions in text items */
function reconstructFractions(items) {
  const merged = [];
  const visited = new Set();

  for (let i = 0; i < items.length; i++) {
    if (visited.has(i)) continue;
    const top = items[i];

    if (!isFractionPart(top.text, top.w)) {
      merged.push(top);
      continue;
    }

    let bestDenIndex = -1;
    let bestDenGap = Infinity;

    for (let j = 0; j < items.length; j++) {
      if (i === j || visited.has(j)) continue;
      const bottom = items[j];

      if (!isFractionPart(bottom.text, bottom.w)) continue;

      const gap = top.y - bottom.y;
      if (gap <= 4 || gap > 28) continue;

      const topCenter = top.x + top.w / 2;
      const bottomCenter = bottom.x + bottom.w / 2;
      const hDiff = Math.abs(topCenter - bottomCenter);

      const overlap = Math.min(top.x + top.w, bottom.x + bottom.w) - Math.max(top.x, bottom.x);
      const minW = Math.min(top.w, bottom.w);

      if (hDiff < 15 && (overlap > 0 || minW < 10)) {
        if (gap < bestDenGap) {
          bestDenGap = gap;
          bestDenIndex = j;
        }
      }
    }

    if (bestDenIndex !== -1) {
      const bottom = items[bestDenIndex];
      visited.add(i);
      visited.add(bestDenIndex);

      // Reconstruct as a LaTeX fraction command: \frac{numerator}{denominator}
      const fractionText = `\\frac{${top.text.trim()}}{${bottom.text.trim()}}`;

      merged.push({
        text: fractionText,
        x: Math.min(top.x, bottom.x),
        y: (top.y + bottom.y) / 2,
        w: Math.max(top.w, bottom.w),
        h: top.y - bottom.y + Math.max(top.h, bottom.h)
      });
    } else {
      merged.push(top);
    }
  }
  return merged;
}

/** Group items on a page into layout-sorted lines */
function groupItemsIntoLines(items) {
  items.sort((a, b) => {
    if (Math.abs(a.y - b.y) > 6) {
      return b.y - a.y;
    }
    return a.x - b.x;
  });

  const lines = [];
  let currentLine = [];
  let currentY = null;

  for (const item of items) {
    if (currentY === null) {
      currentLine.push(item);
      currentY = item.y;
    } else if (Math.abs(item.y - currentY) <= 8) {
      currentLine.push(item);
    } else {
      currentLine.sort((a, b) => a.x - b.x);
      lines.push(currentLine);
      currentLine = [item];
      currentY = item.y;
    }
  }
  if (currentLine.length > 0) {
    currentLine.sort((a, b) => a.x - b.x);
    lines.push(currentLine);
  }

  return lines.map(line => {
    let lineStr = '';
    for (let i = 0; i < line.length; i++) {
      const current = line[i];
      const prev = line[i - 1];
      if (prev) {
        const gap = current.x - (prev.x + prev.w);
        if (gap > 4) {
          lineStr += ' ';
        }
      }
      lineStr += current.text;
    }
    return lineStr;
  });
}

/** Load PDF using PDF.js, reconstruct visual structure, and extract text */
export async function extractTextFromPdf(arrayBuffer) {
  
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = '';

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    
    const items = textContent.items.map(item => ({
      text: item.str,
      x: item.transform[4],
      y: item.transform[5],
      w: item.width || 0,
      h: item.height || 0
    })).filter(item => item.text.trim().length > 0);

    const mergedItems = reconstructFractions(items);
    const lines = groupItemsIntoLines(mergedItems);
    fullText += lines.join('\n') + '\n';
  }
  return fullText;
}

/** Auto-wrap math parts in $$ delimiters and clean math symbols */
export function autoWrapMath(text) {
  let cleaned = text;
  cleaned = cleaned.replace(/√(\d+)/g, '\\sqrt{$1}');
  cleaned = cleaned.replace(/√([a-zA-Z\theta\beta\alpha])/g, '\\sqrt{$1}');
  cleaned = cleaned.replace(/θ/g, '\\theta');
  cleaned = cleaned.replace(/β/g, '\\beta');
  cleaned = cleaned.replace(/α/g, '\\alpha');
  cleaned = cleaned.replace(/°/g, '^{\\circ}');
  
  if (/^\$\$.*\$\$$/.test(cleaned) || /^\$.*\$$/.test(cleaned)) {
    return cleaned;
  }

  const parts = cleaned.split(/([\u0980-\u09ff]+(?:[\s,।\?]+[\u0980-\u09ff]+)*)/);
  
  return parts.map(part => {
    if (!part || /[\u0980-\u09ff]/.test(part)) {
      return part;
    }
    
    const trimmed = part.trim();
    if (!trimmed) return part;
    if (/^[,\.\?\!\-\s]+$/.test(trimmed)) return part;
    
    const isRoman = /^[ivxIVX]+$/.test(trimmed);
    
    let isMath = false;
    if (!isRoman) {
      isMath = /[0-9\+\-\*\/=\^\\_√±θβα°]/.test(trimmed) || 
               /\b(cos|sin|tan|sec|cosec|cot|log|lim|ln)\b/i.test(trimmed) ||
               trimmed.includes('\\frac') ||
               trimmed.includes('\\sqrt') ||
               (trimmed.length <= 4 && /^[a-zA-Z]+$/.test(trimmed));
    }
                   
    if (isMath) {
      const leadSpace = part.startsWith(' ') ? ' ' : '';
      const trailSpace = part.endsWith(' ') ? ' ' : '';
      return `${leadSpace}$$${trimmed}$$${trailSpace}`;
    }
    
    return part;
  }).join('');
}

/** Robust line-by-line MCQ parser */
export function parseMcqText(text) {
  const lines = text.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
  const mcqs = [];
  let currentMcq = null;

  const optARegex = /^[\(]?[Aaক১][\.\)\]\-\s]+/;
  const optBRegex = /^[\(]?[Bbখ২][\.\)\]\-\s]+/;
  const optCRegex = /^[\(]?[Ccগ৩][\.\)\]\-\s]+/;
  const optDRegex = /^[\(]?[Ddঘ৪][\.\)\]\-\s]+/;

  function getOptionLetter(line) {
    if (optARegex.test(line)) return 'A';
    if (optBRegex.test(line)) return 'B';
    if (optCRegex.test(line)) return 'C';
    if (optDRegex.test(line)) return 'D';
    return null;
  }

  function parseInlineOptions(line) {
    const regex = /([\(]?([A-Da-dক-ঘ১-৪])[\.\)\]\-\s]+)/g;
    const matches = [];
    let match;
    while ((match = regex.exec(line)) !== null) {
      matches.push({
        fullMarker: match[1],
        letter: match[2].toUpperCase(),
        index: match.index
      });
    }

    if (matches.length < 2) return null;

    const parsed = {};
    for (let i = 0; i < matches.length; i++) {
      const current = matches[i];
      const next = matches[i + 1];
      const start = current.index + current.fullMarker.length;
      const end = next ? next.index : line.length;
      const val = line.slice(start, end).trim();
      
      let norm = current.letter;
      if (['ক', '১'].includes(norm)) norm = 'A';
      if (['খ', '২'].includes(norm)) norm = 'B';
      if (['গ', '৩'].includes(norm)) norm = 'C';
      if (['ঘ', '৪'].includes(norm)) norm = 'D';

      if (['A', 'B', 'C', 'D'].includes(norm)) {
        parsed[norm] = val;
      }
    }
    return parsed;
  }

  const ansRegex = /^(?:Correct|Answer|Ans|Correct\s*Option|উত্তর|সঠিক\s*উত্তর|সঠিক|ans|correct)[:\s]+([A-Dক-ঘa-d১-৪])/i;

  function getCorrectLetter(line) {
    const match = line.match(ansRegex);
    if (match) {
      let l = match[1].toUpperCase();
      if (['ক', '১'].includes(l)) return 'A';
      if (['খ', '২'].includes(l)) return 'B';
      if (['গ', '৩'].includes(l)) return 'C';
      if (['ঘ', '৪'].includes(l)) return 'D';
      return l;
    }
    return null;
  }

  function finalizeMcq(mcq) {
    if (!mcq) return null;
    return {
      ...mcq,
      question: autoWrapMath(mcq.question),
      optionA: autoWrapMath(mcq.optionA),
      optionB: autoWrapMath(mcq.optionB),
      optionC: autoWrapMath(mcq.optionC),
      optionD: autoWrapMath(mcq.optionD)
    };
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const correctLetter = getCorrectLetter(line);
    if (correctLetter && currentMcq) {
      currentMcq.correct = correctLetter;
      continue;
    }

    const inlineOptions = parseInlineOptions(line);
    if (inlineOptions && currentMcq) {
      if (inlineOptions.A) currentMcq.optionA = inlineOptions.A;
      if (inlineOptions.B) currentMcq.optionB = inlineOptions.B;
      if (inlineOptions.C) currentMcq.optionC = inlineOptions.C;
      if (inlineOptions.D) currentMcq.optionD = inlineOptions.D;
      continue;
    }

    const singleOptLetter = getOptionLetter(line);
    if (singleOptLetter && currentMcq) {
      let val = line;
      if (singleOptLetter === 'A') val = line.replace(optARegex, '').trim();
      if (singleOptLetter === 'B') val = line.replace(optBRegex, '').trim();
      if (singleOptLetter === 'C') val = line.replace(optCRegex, '').trim();
      if (singleOptLetter === 'D') val = line.replace(optDRegex, '').trim();

      currentMcq[`option${singleOptLetter}`] = val;
      continue;
    }

    const isNumberedQuestion = /^(?:Question\s*\d+|[Qq](?:\.|uestion)?\s*\d+|[\d১-৯০]+[\.\)\]।]+|\([\d১-৯০]+\)|\[[\d১-৯০]+\])/.test(line);

    if (isNumberedQuestion || !currentMcq) {
      if (currentMcq) {
        mcqs.push(finalizeMcq(currentMcq));
      }
      
      let qText = line.replace(/^(?:Question\s*\d+|[Qq](?:\.|uestion)?\s*\d+|[\d১-৯০]+[\.\)\]।]+|\([\d১-৯০]+\)|\[[\d১-৯০]+\])/, '').trim();
      // Remove leading colon/Dari/etc. if present
      qText = qText.replace(/^[:।\s]+/, '').trim();

      currentMcq = {
        ...blankMcq(),
        question: qText,
        optionA: '',
        optionB: '',
        optionC: '',
        optionD: '',
        correct: 'A'
      };
    } else {
      if (currentMcq.optionA === '' && currentMcq.optionB === '' && currentMcq.optionC === '' && currentMcq.optionD === '') {
        currentMcq.question += ' ' + line;
      }
    }
  }

  if (currentMcq) {
    mcqs.push(finalizeMcq(currentMcq));
  }

  return mcqs;
}
