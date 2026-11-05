"""Typographic normalizer applied to every chapter (1-13) before it is published or bound.

Fixes:
1. Body paragraphs are justified (both edges), as in a printed academic book.
2. Mixed Arabic/Latin text keeps its order: an RLM mark closes every Latin run that sits
   inside an Arabic paragraph, so trailing years, parentheses and punctuation stay with
   the English words instead of jumping across the line.
3. Lists are indented from the start (right) edge with a hanging bullet/number.
"""
import re
import sys

import docx
from docx.oxml import parse_xml

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RLM = "‏"
ARABIC = re.compile(r"[؀-ۿ]")


def _ptext(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def _is_latin_run(r):
    rtl = r.find(f"{W}rPr/{W}rtl")
    return rtl is not None and rtl.get(W + "val") == "false"


def fix_paragraph(p, in_table=False):
    ppr = p.find(W + "pPr")
    if ppr is None or ppr.find(W + "bidi") is None:
        return
    text = _ptext(p)
    if not ARABIC.search(text):
        return
    # 2. close Latin runs with RLM so neutrals resolve to the Arabic side
    runs = p.findall(f".//{W}r")
    for i, r in enumerate(runs):
        if not _is_latin_run(r):
            continue
        ts = r.findall(W + "t")
        if not ts:
            continue
        last = (ts[-1].text or "").rstrip()
        nxt = "".join(t.text or "" for t in runs[i + 1].findall(W + "t")) if i + 1 < len(runs) else ""
        # only a closing bracket left between English and Arabic text can drift across the line
        if last.endswith((")", "]")) and nxt[:1] and not re.match(r"[A-Za-z0-9.,%]", nxt[:1]) \
                and not (ts[-1].text or "").endswith(RLM):
            ts[-1].text = (ts[-1].text or "") + RLM
    # 1. justify running text (not headings, captions, centered lines, table cells)
    if in_table:
        return
    style = ppr.find(W + "pStyle")
    sval = style.get(W + "val") if style is not None else ""
    jc = ppr.find(W + "jc")
    jval = jc.get(W + "val") if jc is not None else "start"
    if sval in ("Heading1", "Heading2", "BookChapter", "BookPart"):
        return
    if jval in ("center",):
        return
    if len(text) < 90 and sval != "ListParagraph":
        return  # short labels and one-line items stay start-aligned
    # LibreOffice can loop forever laying out a justified bidi line that holds several separate
    # number runs (e.g. "50% ... 20% ... 20%"); such paragraphs stay start-aligned.
    if sum(1 for r in runs if _is_latin_run(r)) >= 3 and len(text) < 400:
        return
    if jc is None:
        jc = parse_xml(f'<w:jc {W_NS} w:val="both"/>')
        ppr.append(jc)
    else:
        jc.set(W + "val", "both")


def fix_numbering(doc):
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        return
    for lvl in numbering.iter(W + "lvl"):
        ilvl = int(lvl.get(W + "ilvl", "0"))
        ppr = lvl.find(W + "pPr")
        if ppr is None:
            continue
        ind = ppr.find(W + "ind")
        if ind is None:
            continue
        start = 720 + 360 * ilvl
        for a in list(ind.attrib):
            del ind.attrib[a]
        # in a bidi paragraph both w:start (Word) and w:left (LibreOffice) mean the leading (right) edge
        ind.set(W + "start", str(start))
        ind.set(W + "left", str(start))
        ind.set(W + "hanging", "360")


def normalize_document(doc):
    body = doc.element.body
    for p in body.iter(W + "p"):
        in_table = any(a.tag == W + "tc" for a in p.iterancestors())
        fix_paragraph(p, in_table)
    fix_numbering(doc)
    return doc


def normalize_file(src, dst=None):
    d = docx.Document(src)
    normalize_document(d)
    d.save(dst or src)


if __name__ == "__main__":
    normalize_file(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
