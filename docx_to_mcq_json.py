#!/usr/bin/env python3
"""
docx_to_mcq_json.py
A universal script to convert Bengali MCQ .docx to JSON.
Handles both multi-line and inline options, mixed text layouts, and raw text formulas.
"""
import json
import re
import zipfile
from pathlib import Path


# ─────────────────────── OMML → LaTeX ────────────────────────────
def _clean(s):
    return re.sub(r"<[^>]+>", "", s).strip()


def omml_to_latex(xml):
    s = xml
    meta = [
        "m:rPr", "m:sPr", "m:fPr", "m:radPr", "m:dPr", "m:naryPr",
        "m:accPr", "m:groupChrPr", "m:barPr", "m:phantPr", "w:rPr", "w:pPr",
    ]
    for tag in meta:
        s = re.sub(rf"<{re.escape(tag)}(?:\s[^>]*)?>[\s\S]*?</{re.escape(tag)}>", "", s)
        s = re.sub(rf"<{re.escape(tag)}(?:\s[^>]*)?/>", "", s)

    def _run(m):
        t = re.search(r"<(?:m:t|w:t)(?:\s[^>]*)?>([^<]*)</(?:m:t|w:t)>", m.group(1))
        return t.group(1) if t else ""

    s = re.sub(r"<(?:m:r|w:r)(?:\s[^>]*)?>[\s\S]*?</(?:m:r|w:r)>", _run, s)
    prev = None
    
    for _ in range(60):
        if prev == s:
            break
        prev = s

        def _frac(m):
            return rf"\frac{{{_clean(m.group(1))}}}{{{_clean(m.group(2))}}}"
        s = re.sub(
            r"<m:f(?:\s[^>]*)?>(?:<[^>]+>)*<m:num(?:\s[^>]*)?>(?P<n>[\s\S]*?)</m:num>(?:<[^>]+>)*<m:den(?:\s[^>]*)?>(?P<d>[\s\S]*?)</m:den>[\s\S]*?</m:f>",
            _frac, s, count=1)

        def _rad(m):
            d, r = _clean(m.group(1)), _clean(m.group(2))
            return rf"\sqrt[{d}]{{{r}}}" if d else rf"\sqrt{{{r}}}"
        s = re.sub(
            r"<m:rad(?:\s[^>]*)?>(?:<[^>]+>)*<m:deg(?:\s[^>]*)?>(?P<d>[\s\S]*?)</m:deg>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>(?P<e>[\s\S]*?)</m:e>[\s\S]*?</m:rad>",
            _rad, s, count=1)

        def _sup(m):
            return f"{_clean(m.group(1))}^{{{_clean(m.group(2))}}}"
        s = re.sub(
            r"<m:sSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>(?P<e>[\s\S]*?)</m:e>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>(?P<s>[\s\S]*?)</m:sup>[\s\S]*?</m:sSup>",
            _sup, s, count=1)

        def _sub(m):
            return f"{_clean(m.group(1))}_{{{_clean(m.group(2))}}}"
        s = re.sub(
            r"<m:sSub(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>(?P<e>[\s\S]*?)</m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>(?P<s>[\s\S]*?)</m:sub>[\s\S]*?</m:sSub>",
            _sub, s, count=1)

        def _subsup(m):
            b, sb, sp = _clean(m.group(1)), _clean(m.group(2)), _clean(m.group(3))
            return f"{b}_{{{sb}}}^{{{sp}}}"
        s = re.sub(
            r"<m:sSubSup(?:\s[^>]*)?>(?:<[^>]+>)*<m:e(?:\s[^>]*)?>(?P<e>[\s\S]*?)</m:e>(?:<[^>]+>)*<m:sub(?:\s[^>]*)?>(?P<sb>[\s\S]*?)</m:sub>(?:<[^>]+>)*<m:sup(?:\s[^>]*)?>(?P<sp>[\s\S]*?)</m:sup>[\s\S]*?</m:sSubSup>",
            _subsup, s, count=1)

        def _delim(m):
            e = re.search(r"<m:d(?:\s[^>]*)>([\s\S]*?)</m:e>", m.group(1))
            return f"({_clean(e.group(1) if e else m.group(1))})"
        s = re.sub(r"<m:d(?:\s[^>]*)>([\s\S]*?)</m:d>", _delim, s, count=1)

        for tag in ("e", "num", "den", "sub", "sup", "deg", "oMath", "oMathPara", "mr"):
            s = re.sub(rf"<m:{tag}(?:\s[^>]*)>([\s\S]*?)</m:{tag}>", r"\1", s)
        s = re.sub(r"<m:[a-zA-Z]+(?:\s[^>]*)?>[\s\S]*?</m:[a-zA-Z]+>",
                   lambda m: _clean(m.group(0)), s)
    return _clean(s)


# ─────────────────── Paragraph text extraction ───────────────────
def get_paragraph_text(p_xml):
    text = ""
    for m in re.finditer(
        r"<(w:r|m:oMath|m:oMathPara)(?:\s[^>]*)>([\s\S]*?)</\1>", p_xml
    ):
        tag, content = m.group(1), m.group(2)
        if tag == "w:r":
            if re.search(r"<w:br\b", content):
                text += "\n"
            for tm in re.finditer(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", content):
                text += tm.group(1)
        elif tag == "m:oMath":
            lt = omml_to_latex(content)
            if lt:
                text += f" ${lt}$ "
        elif tag == "m:oMathPara":
            for om in re.finditer(r"<m:oMath(?:\s[^>]*)>([\s\S]*?)</m:oMath>", content):
                lt = omml_to_latex(om.group(1))
                if lt:
                    text += f" $${lt}$$ "