'use strict';
// gemini-extract.js — turns raw MCQ text (Bengali/English/mixed, with math)
// into a clean array of MCQ objects using Google Gemini (free tier).
//
// Env vars:
//   GEMINI_API_KEY  (required)  — from https://aistudio.google.com/apikey
//   GEMINI_MODEL    (optional)  — default "gemini-2.5-flash"
//
// Returns: [{ passage, question, optionA, optionB, optionC, optionD, correct }]

// Fallback chain: try each model in order; only move to the next if the
// previous one FAILS (rate-limit / error). Override with env GEMINI_MODELS
// (comma-separated) or a single GEMINI_MODEL.
const DEFAULT_MODELS = (
  process.env.GEMINI_MODELS ||
  process.env.GEMINI_MODEL ||
  'gemini-2.5-flash,gemini-2.5-flash-lite,gemini-3-flash-preview'
)
  .split(',')
  .map((m) => m.trim())
  .filter(Boolean);

const DEFAULT_MODEL = DEFAULT_MODELS[0];

const RESPONSE_SCHEMA = {
  type: 'ARRAY',
  items: {
    type: 'OBJECT',
    properties: {
      passage:  { type: 'STRING' },
      question: { type: 'STRING' },
      optionA:  { type: 'STRING' },
      optionB:  { type: 'STRING' },
      optionC:  { type: 'STRING' },
      optionD:  { type: 'STRING' },
      correct:  { type: 'STRING', enum: ['A', 'B', 'C', 'D'] },
    },
    required: ['question', 'optionA', 'optionB', 'optionC', 'optionD', 'correct'],
    propertyOrdering: ['passage', 'question', 'optionA', 'optionB', 'optionC', 'optionD', 'correct'],
  },
};

const SYSTEM_PROMPT = [
  'You extract multiple-choice questions (MCQs) from raw exam text.',
  'The text may be in Bengali, English, or a mix, and may contain math (fractions like 3/5, roots like √10, powers like p^2).',
  '',
  'Rules:',
  '1. Extract EVERY MCQ you find. Do not skip any and do not invent any.',
  '2. Each MCQ has exactly four options mapped to A, B, C, D in the order they appear.',
  '   - If the source labels options with Bengali letters (ক/খ/গ/ঘ) or digits (১/২/৩/৪), map them to A/B/C/D in order.',
  '   - Put ONLY the option content in optionA..optionD. Strip the leading label (e.g. "A.", "ক)", "১।").',
  '3. "correct" is the letter (A/B/C/D) of the correct answer.',
  '   - Find it from lines like "উত্তর", "সঠিক উত্তর", "Answer", "Ans", "Correct".',
  '   - If the answer is given as ক/খ/গ/ঘ or ১/২/৩/৪, map it to A/B/C/D.',
  '   - If no answer is stated for a question, set correct to "A".',
  '4. "passage": stimulus / উদ্দীপক / অনুচ্ছেদ / চিত্র handling.',
  '   - A directive line like "নিচের উদ্দীপকটি পড়ে ৪ ও ৫ নম্বর প্রশ্নের উত্তর দাও" introduces a shared stimulus.',
  '   - Copy that directive line AND the stimulus text/figure/table that follows it into the "passage" field of EVERY question the directive covers (e.g. both Q4 and Q5, or Q5/Q6/Q7 for a three-question set).',
  '   - Do not drop these directive lines. If a question has no stimulus, use an empty string.',
  '5. Math placeholders: the text may contain opaque tokens like \u3016M0\u3017, \u3016M1\u3017 that stand in for math. Copy each token EXACTLY as-is into whatever field it belongs to. Never modify, translate, split, renumber, or drop a token.',
  '6. Do NOT include the question number in the question text.',
  '',
  'Return ONLY the JSON array described by the schema. No commentary.',
].join('\n');

function normalizeMcqs(arr) {
  const out = [];
  const S = (v) => (v == null ? '' : String(v)).trim();
  for (const m of Array.isArray(arr) ? arr : []) {
    if (!m || typeof m !== 'object') continue;
    const q = S(m.question);
    const A = S(m.optionA), B = S(m.optionB), C = S(m.optionC), D = S(m.optionD);
    let correct = S(m.correct).toUpperCase().charAt(0);
    if (!['A', 'B', 'C', 'D'].includes(correct)) correct = 'A';
    if (q && (A || B || C || D)) {
      out.push({ passage: S(m.passage), question: q, optionA: A, optionB: B, optionC: C, optionD: D, correct });
    }
  }
  return out;
}

// Exposed so callers can decide whether to even attempt Gemini.
function hasGeminiKey() {
  return !!(process.env.GEMINI_API_KEY && process.env.GEMINI_API_KEY.trim());
}

// Try one model. Throws on any failure so the caller can fall through.
async function callGeminiModel(rawText, apiKey, model) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const body = {
    systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] },
    contents: [{ role: 'user', parts: [{ text: rawText }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: RESPONSE_SCHEMA,
      temperature: 0,
      maxOutputTokens: 65536,
    },
  };

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    throw new Error(`Gemini API ${resp.status}: ${errText.slice(0, 300)}`);
  }

  const data = await resp.json();
  const cand = data && data.candidates && data.candidates[0];
  if (!cand) throw new Error('Gemini returned no candidates' + (data && data.promptFeedback ? ' (blocked: ' + JSON.stringify(data.promptFeedback) + ')' : ''));

  const parts = (cand.content && cand.content.parts) || [];
  let text = parts.map((p) => p.text || '').join('').trim();
  if (!text) throw new Error('Gemini returned empty text');

  // responseMimeType=json should give raw JSON, but strip fences just in case
  text = text.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error('Could not parse Gemini JSON: ' + e.message);
  }
  const mcqs = normalizeMcqs(parsed);
  if (!mcqs.length) throw new Error('Gemini found no MCQs');
  return mcqs;
}

// Shield $$...$$ math from the LLM: replace each block with an opaque token
// the model carries verbatim, then restore the exact LaTeX afterwards. This
// stops the model from mangling fragile things like matrix row separators (\\).
function protectMath(text) {
  const map = [];
  const protectedText = text.replace(/\$\$[\s\S]*?\$\$|\$[^$\n]*?\$/g, (m) => {
    const token = `\u3016M${map.length}\u3017`; // 〖M0〗 — unlikely to appear naturally
    map.push(m);
    return token;
  });
  return { text: protectedText, map };
}

function restoreMath(str, map) {
  if (!str || !map.length) return str;
  return str.replace(/\u3016M(\d+)\u3017/g, (whole, i) => {
    const v = map[+i];
    return v == null ? whole : v;
  });
}

function restoreMathInMcqs(mcqs, map) {
  if (!map.length) return mcqs;
  for (const m of mcqs) {
    m.passage = restoreMath(m.passage, map);
    m.question = restoreMath(m.question, map);
    m.optionA = restoreMath(m.optionA, map);
    m.optionB = restoreMath(m.optionB, map);
    m.optionC = restoreMath(m.optionC, map);
    m.optionD = restoreMath(m.optionD, map);
  }
  return mcqs;
}

// Public entry: run the fallback chain. Uses each model only if the previous
// one failed. Returns MCQs from the first model that succeeds.
async function extractMcqsWithGemini(rawText, opts = {}) {
  const apiKey = opts.apiKey || process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error('GEMINI_API_KEY is not set');
  if (!rawText || !rawText.trim()) throw new Error('Empty text passed to Gemini');

  const models = opts.models || (opts.model ? [opts.model] : DEFAULT_MODELS);

  // Hide math from the model so it can't corrupt LaTeX (e.g. matrix \\ breaks).
  const { text: safeText, map } = protectMath(rawText);

  let lastErr;
  for (const model of models) {
    try {
      const mcqs = await callGeminiModel(safeText, apiKey, model);
      return restoreMathInMcqs(mcqs, map);
    } catch (err) {
      lastErr = err;
      console.error(`Gemini model "${model}" failed: ${err.message} — trying next model`);
    }
  }
  throw new Error(`All Gemini models failed. Last error: ${lastErr ? lastErr.message : 'unknown'}`);
}

// Build the plain-text prompt body from extracted paragraph lines.
// A blank line is inserted before every numbered paragraph so question
// boundaries stay clear even when Word auto-numbering isn't in the text.
function lineDataToText(lineData) {
  return lineData
    .map((o) => (o.isNumbered ? '\n' : '') + (o.text || ''))
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

module.exports = { extractMcqsWithGemini, normalizeMcqs, lineDataToText, hasGeminiKey, DEFAULT_MODEL };
