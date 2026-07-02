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
      const e = inner.match(/<m:e(?:\s[^>]*)?>([\s\S]*?)<\/m:e>/);
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
  return lines.map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

// ─────────────────── Smart MCQ Parser Engine ──────────────────
