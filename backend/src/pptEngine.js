/**
 * pptEngine.js
 * ------------------------------------------------------------------
 * Generates an MCQ .pptx by REUSING the uploaded master template
 * VERBATIM. We never recreate the layout or redraw shapes. We open the
 * real .pptx (a zip of XML), clone the two canonical slides (a question
 * slide + its answer slide) once per MCQ, and only:
 *   - replace text inside specific shapes (by shape id)
 *   - flip the correct option box fill from pink -> blue on the answer slide
 *   - (optionally) swap the master background image with the course icon
 *
 * All colors, fonts, geometry, shadows, rounded corners, spacing and
 * decorative elements come straight from the template and are untouched.
 *
 * Mechanism is data-driven via TEMPLATE_MAP so that if you replace the
 * template file later, you usually only adjust this map — not the logic.
 */

const JSZip = require('jszip');
let temml = null;
let mml2omml = null;

async function loadMathLibraries() {
  if (!temml) {
    const mod = await import('temml');
    temml = mod.default || mod;
  }
  if (!mml2omml) {
    const mod = await import('mathml2omml');
    mml2omml = mod.mml2omml || mod.default?.mml2omml || mod.default;
  }
}

/* ------------------------------------------------------------------ *
 * TEMPLATE CONTRACT
 * Decoded from Civics2nd_Paper_Final_Touch.pptx.
 *   - slide1.xml = canonical QUESTION slide
 *   - slide2.xml = canonical ANSWER slide (same as Q + one box blued)
 *   - shapes are addressed by their <p:cNvPr id="..."> value
 * ------------------------------------------------------------------ */
const TEMPLATE = {
  questionSlide: 'ppt/slides/slide1.xml',
  answerSlide: 'ppt/slides/slide2.xml',
  // master background image to replace with the course icon
  iconMediaPath: 'ppt/media/image1.png',
  // shape ids -> logical field
  shapeIds: {
    mcqNumber: '202',
    question: '203',
    options: { A: '206', B: '209', C: '212', D: '215' }, // value boxes
  },
  // The icon box and the course-mnemonic box live on the LAYOUT
  // (slideLayout7.xml), so editing them once updates every slide.
  layout: {
    path: 'ppt/slideLayouts/slideLayout7.xml',
    relsPath: 'ppt/slideLayouts/_rels/slideLayout7.xml.rels',
    iconBoxId: '2', // top-left rounded box ("Icon") — image fills this box
    mnemonicBoxId: '6', // top-right rounded box ("Course Mnemonic") text
  },
  // option letters are fixed Bengali glyphs in the template (ক/খ/গ/ঘ);
  // we do NOT touch the letter shapes, only the value boxes.
  colors: {
    pink: 'E2098D', // unselected option fill
    blue: '002060', // correct option fill (matches question/letter blue)
  },
};

/* ----------------------------- helpers ---------------------------- */

function getScaleFactor(len, isQuestion) {
  if (isQuestion) {
    if (len <= 100) return 1.0;
    if (len <= 200) return 0.85;
    if (len <= 350) return 0.70;
    if (len <= 500) return 0.55;
    return 0.45; // Minimum scale factor (e.g. 9pt if baseline is 20pt)
  } else {
    // Options
    if (len <= 25) return 1.0;
    if (len <= 50) return 0.85;
    if (len <= 80) return 0.70;
    return 0.60;
  }
}

function adjustFontSize(rPr, textLength, isQuestion) {
  if (!rPr) {
    const baseSz = isQuestion ? 2000 : 1800;
    const scaled = Math.round(baseSz * getScaleFactor(textLength, isQuestion));
    return `<a:rPr sz="${scaled}"/>`;
  }
  
  // Find current size (in hundredths of a point)
  const szMatch = rPr.match(/sz="(\d+)"/);
  if (!szMatch) {
    const baseSz = isQuestion ? 2000 : 1800;
    const scaled = Math.round(baseSz * getScaleFactor(textLength, isQuestion));
    if (rPr.startsWith('<a:rPr')) {
      return rPr.replace('<a:rPr', `<a:rPr sz="${scaled}"`);
    } else {
      return `<a:rPr sz="${scaled}"/>`;
    }
  }

  const currentSz = parseInt(szMatch[1], 10);
  const factor = getScaleFactor(textLength, isQuestion);
  const newSz = Math.round(currentSz * factor);
  
  return rPr.replace(/sz="\d+"/, `sz="${newSz}"`);
}

/**
 * Force "shrink text on overflow" (PowerPoint AutoFit) on a shape's text body,
 * so a long question or long option automatically scales down to fit its fixed
 * box instead of spilling out. Replaces any noAutofit/spAutoFit with normAutofit.
 */
function ensureNormAutofit(txBody) {
  // Already has normAutofit → leave it (PowerPoint recomputes the scale).
  if (/<a:normAutofit\b/.test(txBody)) return txBody;
  // Swap an explicit no-autofit / shape-autofit for normAutofit.
  if (/<a:noAutofit\s*\/>/.test(txBody)) return txBody.replace(/<a:noAutofit\s*\/>/, '<a:normAutofit/>');
  if (/<a:spAutoFit\s*\/>/.test(txBody)) return txBody.replace(/<a:spAutoFit\s*\/>/, '<a:normAutofit/>');
  // Self-closing bodyPr → expand and add normAutofit.
  if (/<a:bodyPr\b[^>]*\/>/.test(txBody)) {
    return txBody.replace(/<a:bodyPr\b([^>]*)\/>/, '<a:bodyPr$1><a:normAutofit/></a:bodyPr>');
  }
  // Open/close bodyPr with no autofit child → insert one.
  if (/<a:bodyPr\b[^>]*>[\s\S]*?<\/a:bodyPr>/.test(txBody)) {
    return txBody.replace(/(<a:bodyPr\b[^>]*>)/, '$1<a:normAutofit/>');
  }
  return txBody;
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Find the exact [start,end) span of the <p:sp>...</p:sp> element that
 * CONTAINS a <p:cNvPr id="shapeId">. Shapes can be nested inside <p:grpSp>,
 * so we cannot use a naive regex — we anchor on the cNvPr, walk left to the
 * nearest enclosing <p:sp>, then bracket-match forward to its </p:sp>.
 * Returns {start, end, xml} or null.
 */
function findShapeSpan(xml, shapeId) {
  const anchor = xml.indexOf(`<p:cNvPr id="${shapeId}"`);
  if (anchor === -1) return null;

  // walk left to the nearest "<p:sp>" that opens before the anchor
  const start = xml.lastIndexOf('<p:sp>', anchor);
  if (start === -1) return null;

  // bracket-match <p:sp> ... </p:sp> from start, accounting for nesting
  let depth = 0;
  let i = start;
  const open = '<p:sp>';
  const close = '</p:sp>';
  while (i < xml.length) {
    const nextOpen = xml.indexOf(open, i);
    const nextClose = xml.indexOf(close, i);
    if (nextClose === -1) return null;
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++;
      i = nextOpen + open.length;
    } else {
      depth--;
      i = nextClose + close.length;
      if (depth === 0) {
        return { start, end: i, xml: xml.slice(start, i) };
      }
    }
  }
  return null;
}

/**
 * Extract the complete <a:rPr ...>...</a:rPr> OR self-closed <a:rPr .../>
 * from the start of a run string, handling nested self-closing children.
 * Returns the rPr substring (or '' if none).
 */
function extractRPr(runXml) {
  const open = runXml.indexOf('<a:rPr');
  if (open === -1) return '';
  // find end of the opening <a:rPr ...>  (could be '/>' or '>')
  let i = open + 6;
  let selfClose = false;
  while (i < runXml.length) {
    const ch = runXml[i];
    if (ch === '>') {
      selfClose = runXml[i - 1] === '/';
      i++;
      break;
    }
    i++;
  }
  if (selfClose) return runXml.slice(open, i); // <a:rPr .../>
  // otherwise find matching </a:rPr>
  const close = runXml.indexOf('</a:rPr>', i);
  if (close === -1) return runXml.slice(open, i); // fallback
  return runXml.slice(open, close + '</a:rPr>'.length);
}

/**
 * Helper to split text into runs (standard text) and math equations (OMML).
 * Compiles LaTeX formulas enclosed in $$...$$ or $...$ into OMML using temml and mathml2omml.
 * Also converts \n newlines into <a:br/> OOXML line-break elements.
 */
function generateRunsAndMath(newText, rPr) {
  // Split on \n first so each line is processed independently.
  // Lines are joined with <a:br> (soft line break) so PowerPoint renders them correctly.
  const lines = newText.split('\n');
  if (lines.length > 1) {
    return lines
      .map((line) => generateLineRuns(line, rPr))
      .join(`<a:br>${rPr || '<a:rPr/>'}</a:br>`);
  }
  return generateLineRuns(newText, rPr);
}

/** Generate runs (text + inline math) for a single line (no \n). */
function generateLineRuns(newText, rPr) {
  const regex = /(\$\$?)([^\$]+?)(\$\$?)/g;
  let match;
  let lastIndex = 0;
  const tokens = [];

  while ((match = regex.exec(newText)) !== null) {
    if (match.index > lastIndex) {
      tokens.push({ type: 'text', value: newText.slice(lastIndex, match.index) });
    }
    tokens.push({ type: 'math', value: match[2].trim(), raw: match[0] });
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < newText.length) {
    tokens.push({ type: 'text', value: newText.slice(lastIndex) });
  }

  if (tokens.length === 0) {
    return `<a:r>${rPr}<a:t></a:t></a:r>`;
  }

  let outXml = '';
  for (const token of tokens) {
    if (token.type === 'text') {
      outXml += `<a:r>${rPr}<a:t>${escapeXml(token.value)}</a:t></a:r>`;
    } else {
      try {
        if (!temml || !mml2omml) {
          throw new Error('Math libraries not loaded');
        }
        const mathml = temml.renderToString(token.value, { displayMode: false });
        let omml = mml2omml(mathml);
        // Inject rPr into all math runs to inherit font size, color, family, etc.
        omml = omml.replace(/<m:r>/g, `<m:r>${rPr}`);
        // Embed the OMML in an a14:m drawing extension container
        outXml += `<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${omml}</m:oMathPara></a14:m>`;
      } catch (err) {
        console.error('Math rendering failed for:', token.value, err);
        // Fallback: render the raw LaTeX string as standard text
        outXml += `<a:r>${rPr}<a:t>${escapeXml(token.raw)}</a:t></a:r>`;
      }
    }
  }
  return outXml;
}

/**
 * Replace the visible text of a shape identified by cNvPr id.
 * Strategy: within the shape's first <a:p> that contains a run, replace the
 * FIRST run with the formatted runs and math blocks, and drop every other run in the
 * whole txBody. This preserves the first run's formatting (rPr) exactly
 * and never produces malformed XML.
 */
function setShapeText(xml, shapeId, newText) {
  const span = findShapeSpan(xml, shapeId);
  if (!span) return xml; // shape not present on this slide

  let sp = span.xml;
  const txBodyMatch = sp.match(/<p:txBody>[\s\S]*<\/p:txBody>/);
  if (!txBodyMatch) return xml;

  let txBody = txBodyMatch[0];
  const runs = [...txBody.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((x) => x[0]);
  let rPr = '';

  const isQuestion = (shapeId === TEMPLATE.shapeIds.question);
  const isOption = Object.values(TEMPLATE.shapeIds.options).includes(shapeId);

  if (runs.length === 0) {
    // No run: inject a run into the first <a:p>, deriving rPr from endParaRPr
    txBody = txBody.replace(
      /(<a:p>)([\s\S]*?)(<\/a:p>)/,
      (full, open, inner, close) => {
        const epr = inner.match(/<a:endParaRPr[\s\S]*?(?:\/>|<\/a:endParaRPr>)/);
        if (epr) {
          rPr = epr[0]
            .replace('<a:endParaRPr', '<a:rPr')
            .replace('</a:endParaRPr>', '</a:rPr>');
        }
        if (isQuestion || isOption) {
          rPr = adjustFontSize(rPr, newText.length, isQuestion);
        }
        const newFirst = generateRunsAndMath(newText, rPr);
        return `${open}${newFirst}${inner}${close}`;
      }
    );
  } else {
    // Keep first run's COMPLETE rPr, generate text/math runs; drop all other runs.
    const firstRun = runs[0];
    rPr = extractRPr(firstRun);
    if (isQuestion || isOption) {
      rPr = adjustFontSize(rPr, newText.length, isQuestion);
    }
    const newFirst = generateRunsAndMath(newText, rPr);

    // Replace runs by exact substring, first->newFirst, rest->removed.
    // Iterate by position to avoid replacing a duplicate run elsewhere.
    let out = '';
    let cursor = 0;
    const runMatches = [...txBody.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)];
    runMatches.forEach((rm, idx) => {
      out += txBody.slice(cursor, rm.index);
      out += idx === 0 ? newFirst : '';
      cursor = rm.index + rm[0].length;
    });
    out += txBody.slice(cursor);
    txBody = out;
  }

  // Inject defRPr into the paragraph properties (a:pPr) to ensure the fraction lines
  // and other math borders inherit the correct text color.
  if (rPr) {
    const defRPr = rPr
      .replace(/^<a:rPr/, '<a:defRPr')
      .replace(/<\/a:rPr>$/, '</a:defRPr>');

    if (/<a:pPr([^>]*)\/>/.test(txBody)) {
      txBody = txBody.replace(/<a:pPr([^>]*)\/>/, (m, attrs) => {
        return `<a:pPr${attrs}>${defRPr}</a:pPr>`;
      });
    } else if (/<a:pPr([^>]*)>/.test(txBody)) {
      txBody = txBody.replace(/<a:pPr([^>]*)>/, (m, attrs) => {
        return `<a:pPr${attrs}>${defRPr}`;
      });
    } else {
      txBody = txBody.replace(/<a:p>/, `<a:p><a:pPr>${defRPr}</a:pPr>`);
    }
  }

  // Auto-fit: shrink long question / option text to fit its fixed box.
  if (isQuestion || isOption) {
    txBody = ensureNormAutofit(txBody);
  }

  sp = sp.replace(txBodyMatch[0], txBody);
  return xml.slice(0, span.start) + sp + xml.slice(span.end);
}


/**
 * Flip a single option value-box fill from pink to blue (answer slide only).
 * We scope to the target shape and replace the FIRST solidFill srgbClr,
 * which is the box fill (the line/stroke fill comes after and is left alone).
 */
function highlightOption(xml, shapeId, fromHex, toHex) {
  const span = findShapeSpan(xml, shapeId);
  if (!span) return xml;
  let sp = span.xml;
  // Replace only the FIRST solidFill box color within this shape.
  // The fill block may contain whitespace/newlines between tags, so match
  // tolerantly. The box fill appears before the <a:ln> stroke fill.
  sp = sp.replace(
    new RegExp(
      `(<a:solidFill>\\s*<a:srgbClr val=")${fromHex}("\\s*/>)`
    ),
    `$1${toHex}$2`
  );
  return xml.slice(0, span.start) + sp + xml.slice(span.end);
}

/**
 * Fill a shape (identified by cNvPr id) with a picture (blipFill), preserving
 * the shape's exact position, size and rounded-corner geometry. Used to drop
 * the course icon INTO the top-left "Icon" box on the layout.
 *
 * We:
 *  - remove any existing <a:solidFill>/<p:style> visual fill influence by
 *    inserting a <a:blipFill> into <p:spPr> right after the geometry,
 *  - blank the shape's text (the literal "Icon" label),
 *  - return the modified xml plus the relationship id used (caller wires rels).
 *
 * @param {string} xml         layout xml
 * @param {string} shapeId     target shape id
 * @param {string} embedRId    relationship id of the image (e.g. "rId5")
 */
function fillShapeWithImage(xml, shapeId, embedRId) {
  const span = findShapeSpan(xml, shapeId);
  if (!span) return xml;
  let sp = span.xml;

  // 1) Clear the text label inside the shape.
  sp = sp.replace(
    /<p:txBody>[\s\S]*<\/p:txBody>/,
    '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
  );

  // 2) Insert a blipFill into spPr, right after the geometry (prstGeom),
  //    and strip the <p:style> fill so the picture shows cleanly.
  const blipFill =
    `<a:blipFill rotWithShape="1">` +
    `<a:blip r:embed="${embedRId}"/>` +
    `<a:stretch><a:fillRect/></a:stretch>` +
    `</a:blipFill>`;

  // place after </a:prstGeom> (geometry kept => rounded corners preserved)
  sp = sp.replace(/(<\/a:prstGeom>)/, `$1${blipFill}`);

  // remove the style fillRef tint so it doesn't overlay the image
  // (keep lnRef so the border stays; just drop fillRef color influence)
  // Picture fill in spPr already takes precedence over style fillRef, so
  // no further change is strictly needed, but we leave style intact for border.

  return xml.slice(0, span.start) + sp + xml.slice(span.end);
}
function toBengaliDigits(n) {
  const bengaliDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
  return String(n).split('').map(digit => bengaliDigits[parseInt(digit, 10)] || digit).join('');
}

function populateSlide(xml, mcq, number) {
  const ids = TEMPLATE.shapeIds;
  xml = setShapeText(xml, ids.mcqNumber, `MCQ ${number}`);
  
  let qText = mcq.question;
  if (mcq.passage && mcq.passage.trim()) {
    const bnNum = toBengaliDigits(number);
    qText = `${mcq.passage.trim()}\n${bnNum}। ${mcq.question}`;
  }
  
  xml = setShapeText(xml, ids.question, qText);
  xml = setShapeText(xml, ids.options.A, mcq.optionA);
  xml = setShapeText(xml, ids.options.B, mcq.optionB);
  xml = setShapeText(xml, ids.options.C, mcq.optionC);
  xml = setShapeText(xml, ids.options.D, mcq.optionD);
  return xml;
}

/* ------------------ slide & relationship plumbing ----------------- */

/** Read presentation.xml, sldIdLst, and the rels so we can append slides. */
async function loadDeckParts(zip) {
  const presXml = await zip.file('ppt/presentation.xml').async('string');
  const presRels = await zip
    .file('ppt/_rels/presentation.xml.rels')
    .async('string');
  const contentTypes = await zip.file('[Content_Types].xml').async('string');
  return { presXml, presRels, contentTypes };
}

function nextRId(presRels) {
  const ids = [...presRels.matchAll(/Id="rId(\d+)"/g)].map((m) =>
    parseInt(m[1], 10)
  );
  return `rId${Math.max(0, ...ids) + 1}`;
}

function nextSlideNum(zip) {
  const nums = Object.keys(zip.files)
    .map((f) => f.match(/^ppt\/slides\/slide(\d+)\.xml$/))
    .filter(Boolean)
    .map((m) => parseInt(m[1], 10));
  return Math.max(0, ...nums) + 1;
}

/**
 * Append a slide whose XML is `xml`, copying the rels of `sourceRelsPath`
 * (so layout + notes references stay valid), and wiring it into
 * presentation.xml, its rels, and [Content_Types].xml.
 */
async function appendSlide(zip, xml, sourceRelsXml, state) {
  const num = state.slideCounter++;
  const slidePath = `ppt/slides/slide${num}.xml`;
  const relsPath = `ppt/slides/_rels/slide${num}.xml.rels`;

  zip.file(slidePath, xml);
  zip.file(relsPath, sourceRelsXml); // reuse layout/notes rels verbatim

  // presentation rels: add relationship to this slide
  const rId = nextRId(state.presRels);
  state.presRels = state.presRels.replace(
    '</Relationships>',
    `  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${num}.xml"/>\n</Relationships>`
  );
  state.sldIds.push(rId);

  // content types: ensure slide override present
  if (
    !state.contentTypes.includes(`/ppt/slides/slide${num}.xml`)
  ) {
    state.contentTypes = state.contentTypes.replace(
      '</Types>',
      `  <Override PartName="/ppt/slides/slide${num}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n</Types>`
    );
  }
}

/* ----------------------------- main ------------------------------- */

/**
 * @param {Buffer} templateBuffer  the master .pptx
 * @param {Array}  mcqs            [{question,optionA..D,correct:'A'|'B'|'C'|'D'}]
 * @param {Object} [opts]
 * @param {Buffer} [opts.iconBuffer]      PNG bytes for the top-left "Icon" box
 * @param {Buffer} [opts.mnemonicBuffer]  PNG bytes for the top-right box
 * @returns {Promise<Buffer>}      generated .pptx
 */
async function generatePptx(templateBuffer, mcqs, opts = {}) {
  await loadMathLibraries();
  const { iconBuffer = null, mnemonicBuffer = null } = opts;
  const zip = await JSZip.loadAsync(templateBuffer);

  // Canonical slide XML + their rels (from the first MCQ pair in template)
  const qXml = await zip.file(TEMPLATE.questionSlide).async('string');
  const aXml = await zip.file(TEMPLATE.answerSlide).async('string');
  const qRels = await zip
    .file('ppt/slides/_rels/slide1.xml.rels')
    .async('string');
  const aRels = await zip
    .file('ppt/slides/_rels/slide2.xml.rels')
    .async('string');

  const cleanRels = (xml) =>
    xml.replace(
      /<Relationship\s+[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/notesSlide"[^>]*\/>/g,
      ''
    );
  const qRelsCleaned = cleanRels(qRels);
  const aRelsCleaned = cleanRels(aRels);

  /* ---- Edit the LAYOUT once: fill icon box + mnemonic box with images ---- */
  let layoutXml = await zip.file(TEMPLATE.layout.path).async('string');
  let layoutRels = await zip.file(TEMPLATE.layout.relsPath).async('string');

  // helper: add a PNG to media, register a layout relationship, return rId
  const ensurePngContentType = async () => {
    let ct = await zip.file('[Content_Types].xml').async('string');
    if (!/Extension="png"/.test(ct)) {
      ct = ct.replace(
        '</Types>',
        '  <Default Extension="png" ContentType="image/png"/>\n</Types>'
      );
      zip.file('[Content_Types].xml', ct);
    }
  };
  const addLayoutImage = (buffer, fileName) => {
    zip.file(`ppt/media/${fileName}`, buffer);
    const existing = [...layoutRels.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1]);
    const rId = `rId${Math.max(0, ...existing) + 1}`;
    layoutRels = layoutRels.replace(
      '</Relationships>',
      `  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${fileName}"/>\n</Relationships>`
    );
    return rId;
  };

  if (iconBuffer || mnemonicBuffer) await ensurePngContentType();

  // Course icon → top-left box
  if (iconBuffer) {
    const rId = addLayoutImage(iconBuffer, 'courseIcon.png');
    layoutXml = fillShapeWithImage(layoutXml, TEMPLATE.layout.iconBoxId, rId);
  }

  // Course mnemonic IMAGE → top-right box
  if (mnemonicBuffer) {
    const rId = addLayoutImage(mnemonicBuffer, 'courseMnemonic.png');
    layoutXml = fillShapeWithImage(layoutXml, TEMPLATE.layout.mnemonicBoxId, rId);
  }

  zip.file(TEMPLATE.layout.path, layoutXml);
  zip.file(TEMPLATE.layout.relsPath, layoutRels);

  // Wipe the template's example slides from the deck ordering; we rebuild it.
  const parts = await loadDeckParts(zip);

  // Strip ALL existing slide relationships + sldIdLst so we start clean,
  // but keep the slide parts on disk (we overwrite slide1/2 below anyway).
  let presRels = parts.presRels.replace(
    /\s*<Relationship Id="[^"]*"[^>]*relationships\/slide"[^>]*\/>/g,
    ''
  );
  let presXml = parts.presXml;
  // remove existing sldId entries
  presXml = presXml.replace(/<p:sldId[^>]*\/>/g, '');

  // Remove example slide parts (slide1..slideN) so only generated remain.
  Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .forEach((f) => zip.remove(f));
  Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f))
    .forEach((f) => zip.remove(f));

  const state = {
    presRels,
    contentTypes: parts.contentTypes,
    sldIds: [],
    slideCounter: 1,
  };

  // Build two slides per MCQ
  for (let i = 0; i < mcqs.length; i++) {
    const mcq = mcqs[i];
    const n = i + 1;

    // QUESTION slide
    let q = populateSlide(qXml, mcq, n);
    await appendSlide(zip, q, qRelsCleaned, state);

    // ANSWER slide: same content + flip the correct option to blue
    let a = populateSlide(aXml, mcq, n);
    // The template's slide2 already had option B blued; reset every option to
    // pink first, then blue only the correct one — guarantees correctness
    // regardless of which option the template happened to pre-highlight.
    for (const letter of ['A', 'B', 'C', 'D']) {
      a = highlightOption(
        a,
        TEMPLATE.shapeIds.options[letter],
        TEMPLATE.colors.blue,
        TEMPLATE.colors.pink
      );
    }
    a = highlightOption(
      a,
      TEMPLATE.shapeIds.options[mcq.correct],
      TEMPLATE.colors.pink,
      TEMPLATE.colors.blue
    );
    await appendSlide(zip, a, aRelsCleaned, state);
  }

  // Rebuild sldIdLst in presentation.xml
  let minId = 256;
  const sldIdXml = state.sldIds
    .map((rId, idx) => `<p:sldId id="${minId + idx}" r:id="${rId}"/>`)
    .join('');
  if (/<p:sldIdLst\s*\/>/.test(presXml)) {
    presXml = presXml.replace(
      /<p:sldIdLst\s*\/>/,
      `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`
    );
  } else if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(presXml)) {
    presXml = presXml.replace(
      /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`
    );
  } else {
    // insert after sldMasterIdLst
    presXml = presXml.replace(
      /(<\/p:sldMasterIdLst>)/,
      `$1<p:sldIdLst>${sldIdXml}</p:sldIdLst>`
    );
  }

  zip.file('ppt/presentation.xml', presXml);
  zip.file('ppt/_rels/presentation.xml.rels', state.presRels);
  zip.file('[Content_Types].xml', state.contentTypes);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { generatePptx, TEMPLATE };
