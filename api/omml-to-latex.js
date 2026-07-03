'use strict';
// omml-to-latex.js — convert Word math (OMML) to LaTeX and extract paragraph
// text as lines, wrapping every math block in $$...$$ so the pptx engine
// (temml → OMML) renders it properly. Ported from the local backend so the
// Vercel API and the local server behave identically.

function clean(s) {
  return s.replace(/<[^>]+>/g, '').trim();
}

// Convert an OMML fragment (contents of <m:oMath>) → LaTeX.
function ommlToLatex(xml) {
  let s = xml;

  const metaTags = ['m:rPr', 'm:sPr', 'm:fPr', 'm:radPr', 'm:dPr', 'm:naryPr',
    'm:accPr', 'm:groupChrPr', 'm:barPr', 'm:phantPr', 'w:rPr', 'w:pPr'];
  for (const t of metaTags) {
    s = s.replace(new RegExp(`<${t}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${t}>`, 'g'), '');
    s = s.replace(new RegExp(`<${t}(?:\\s[^>]*)?\\/>`, 'g'), '');
  }

  s = s.replace(/<(?:m:r|w:r)(?:\s[^>]*)?>([\s\S]*?)<\/(?:m:r|w:r)>/g, (_, inner) => {
    const tm = inner.match(/<(?:m:t|w:t)(?:\s[^>]*)?>([^<]*)<\/(?:m:t|w:t)>/);
    return tm ? tm[1] : '';
  });

  let prev = '';
  for (let iter = 0; iter < 40 && prev !== s; iter++) {
    prev = s;

    // Fraction
    s = s.replace(
      /<m:f(?:\s[^>]*)?>(?:<[^>]+>)*<m:num(?:\s[^>]*)?>([\s\S]*?)<\/m:num>(?:<[^>]+>)*<m:den(?:\s[^>]*)?>([\s\S]*?)<\/m:den>[\s\S]*?<\/m:f>/,
      (_, n, d) => `\\frac{${clean(n)}}{${clean(d)}}`
    );
    // Radical
    s = s.replace(
      /<m:rad(?:\s[^>]*)?>(?:<[^>]+>)*<m:deg(?:\s[^>]*)?>([\s\S]*?)<\/m:deg>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>[\s\S]*?<\/m:rad>/,
      (_, deg, rad) => {
        const d = clean(deg).trim();
        return d ? `\\sqrt[${d}]{${clean(rad)}}` : `\\sqrt{${clean(rad)}}`;
      }
    );
    // Superscript
    s = s.replace(
      /<m:sSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSup>/,
      (_, base, exp) => `${clean(base)}^{${clean(exp)}}`
    );
    // Subscript
    s = s.replace(
      /<m:sSub(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>[\s\S]*?<\/m:sSub>/,
      (_, base, sub) => `${clean(base)}_{${clean(sub)}}`
    );
    // Sub + Superscript
    s = s.replace(
      /<m:sSubSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>([\s\S]*?)<\/m:sub>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>([\s\S]*?)<\/m:sup>[\s\S]*?<\/m:sSubSup>/,
      (_, base, sub, sup) => `${clean(base)}_{${clean(sub)}}^{${clean(sup)}}`
    );
    // Delimiter
    s = s.replace(
      /<m:d(?:\s[^>]*)?>([\s\S]*?)<\/m:d>/,
      (_, inner) => {
        const eMatch = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/);
        return eMatch ? `(${clean(eMatch[1])})` : `(${clean(inner)})`;
      }
    );
    // Accent / overline
    s = s.replace(/<m:acc(?:\s[^>]*)?>([\s\S]*?)<\/m:acc>/,
      (_, inner) => { const e = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/); return e ? clean(e[1]) : clean(inner); }
    );
    // Structural wrappers
    for (const tag of ['m:e', 'm:num', 'm:den', 'm:sub', 'm:sup', 'm:deg', 'm:oMath', 'm:oMathPara', 'm:mr']) {
      s = s.replace(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>`, 'g'), '$1');
    }
    // Remaining unknown m: tags → keep text
    s = s.replace(/<m:[a-zA-Z]+(?:\s[^>]*)?>[\s\S]*?<\/m:[a-zA-Z]+>/g, (m) => m.replace(/<[^>]+>/g, ''));
  }

  return s.replace(/<[^>]+>/g, '').trim();
}

function decodeXml(s) {
  return s
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)));
}

// Extract paragraph text as lines. Normal text from <w:t>; math blocks
// (<m:oMath>/<m:oMathPara>) become LaTeX wrapped in $$...$$. Splits on <w:br>.
function getParagraphLines(pXml) {
  let currentLine = '';
  const lines = [];
  const xml = pXml.replace(/<w:tab(?:\s[^>]*)?\/?>/g, ' ');

  const childRx = /<(w:r|m:oMath|m:oMathPara)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/g;
  let m;
  while ((m = childRx.exec(xml)) !== null) {
    const tag = m[1];
    const content = m[2];

    if (tag === 'w:r') {
      const brMatches = content.match(/<w:br\b/g);
      const tRx = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let tm, textPart = '';
      while ((tm = tRx.exec(content)) !== null) textPart += decodeXml(tm[1]);
      if (brMatches) {
        if (currentLine.trim() || textPart.trim()) lines.push((currentLine + textPart).trim());
        for (let i = 0; i < brMatches.length - 1; i++) lines.push('');
        currentLine = '';
      } else {
        currentLine += textPart;
      }
    } else if (tag === 'm:oMath') {
      const latex = ommlToLatex(content);
      if (latex) currentLine += ` $$${latex}$$ `;
    } else if (tag === 'm:oMathPara') {
      const oRx = /<m:oMath(?:\s[^>]*)?>([\s\S]*?)<\/m:oMath>/g;
      let om;
      while ((om = oRx.exec(content)) !== null) {
        const latex = ommlToLatex(om[1]);
        if (latex) currentLine += ` $$${latex}$$ `;
      }
    }
  }
  if (currentLine.trim()) lines.push(currentLine.trim());
  return lines.map((l) => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

module.exports = { ommlToLatex, clean, getParagraphLines };
