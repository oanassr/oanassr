"""Read slide-only content (English titles, leads, card layouts, takeaway bars) from a Ch01-09 deck."""
import re

from pptx import Presentation
from pptx.util import Emu

E = 914400
CARD_FILLS = {"0F4C5C", "2C7A7B", "E09F3E", "B5532A", "3D5A80", "6B4E71"}


def _fill(sh):
    try:
        if sh.fill.type == 1:
            return str(sh.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def _geom(sh):
    return (sh.left / E, sh.top / E, sh.width / E, sh.height / E)


def _texts(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            x, y, w, h = _geom(sh)
            size = None
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        size = r.font.size.pt
                        break
                if size:
                    break
            lines = [p.text for p in sh.text_frame.paragraphs if p.text.strip()]
            out.append({"x": x, "y": y, "w": w, "h": h, "text": sh.text_frame.text.strip(), "lines": lines, "size": size})
    return out


def _inside(t, box, pad=0.02):
    x, y, w, h = box
    return t["x"] >= x - pad and t["y"] >= y - pad and t["x"] + t["w"] <= x + w + pad and t["y"] + t["h"] <= y + h + pad


def parse_section_slide(slide, docx_table_header=None):
    texts = [t for t in _texts(slide) if t["y"] > 0.45]  # skip nav tabs
    en = next((t["text"] for t in texts if abs(t["y"] - 1.20) < 0.03), "")
    title = next((t["text"] for t in texts if abs(t["y"] - 0.72) < 0.03), "")
    lead = next((t["text"] for t in texts if abs(t["y"] - 1.59) < 0.03 and t["x"] < 1.0), "")
    take = None
    shapes = list(slide.shapes)
    for sh in shapes:
        if _fill(sh) == "F1E9DA" and sh.top / E > 4.3:
            box = _geom(sh)
            inner = [t for t in texts if _inside(t, box)]
            if inner:
                take = inner[0]["text"]
    kind, data = None, None
    tbl = next((sh for sh in shapes if sh.has_table), None)
    if tbl is not None:
        rows = [[c.text.strip() for c in r.cells] for r in tbl.table.rows]
        if docx_table_header and rows and rows[0][::-1] == docx_table_header:
            rows = [r[::-1] for r in rows]
        elif rows and rows[0] != (docx_table_header or rows[0]) and rows[0][::-1] != (docx_table_header or rows[0]):
            # different table on the slide: assume RTL order stored left-to-right
            rows = [r[::-1] for r in rows]
        widths = [c.width for c in tbl.table.columns]
        if docx_table_header and rows[0] == docx_table_header and len(widths) == len(rows[0]):
            widths = widths[::-1] if rows[0] != [c.text.strip() for c in tbl.table.rows[0].cells] else widths
        kind, data = "table", {"header": rows[0], "rows": rows[1:], "widths": [w / E for w in widths]}
    else:
        heads = [sh for sh in shapes if _fill(sh) in ("0F4C5C", "B5532A") and abs(sh.height / E - 0.50) < 0.02 and abs(sh.width / E - 4.50) < 0.05]
        if len(heads) == 2:
            groups = []
            for hd in sorted(heads, key=lambda s: -s.left):
                hx, hy, hw, hh = _geom(hd)
                ht = next(t["text"] for t in texts if _inside(t, (hx, hy, hw, hh)))
                body = next((t for t in texts if t["x"] >= hx - 0.05 and t["x"] + t["w"] <= hx + hw + 0.05 and t["y"] > hy + 0.3), None)
                groups.append((ht, body["lines"] if body else []))
            kind, data = "two", groups
        else:
            cards = [sh for sh in shapes if _fill(sh) in CARD_FILLS and sh.height / E > 0.9 and sh.width / E < 4.4]
            if cards:
                items = []
                for c in cards:
                    box = _geom(c)
                    inner = [t for t in texts if _inside(t, box)]
                    num = next((t for t in inner if re.fullmatch(r"\d+", t["text"])), None)
                    rest = [t for t in inner if t is not num]
                    rest.sort(key=lambda t: t["y"])
                    ar = rest[0]["text"] if rest else ""
                    enl = rest[1]["text"] if len(rest) > 1 else ""
                    desc = rest[2]["text"] if len(rest) > 2 else ""
                    items.append((int(num["text"]) if num else 99, ar, enl, desc))
                items.sort()
                tri = any('prst="triangle"' in sh._element.xml for sh in shapes)
                kind = "steps" if tri else "grid"
                data = [(a, e, d) for _, a, e, d in items]
    return {"title": title, "en": en, "lead": lead, "take": take, "kind": kind, "data": data}


def parse_deck(path, ch):
    prs = Presentation(path)
    slides = list(prs.slides)
    cover = _texts(slides[0])
    title_en = next(t["text"] for t in cover if abs(t["y"] - 2.35) < 0.05)
    note = ""
    if slides[0].has_notes_slide:
        note = slides[0].notes_slide.notes_text_frame.text.strip()
    opening = _texts(slides[2])
    oc_en = next(t["text"] for t in opening if abs(t["y"] - 1.20) < 0.03)
    body = sorted([t for t in opening if t["x"] > 3.9 and t["y"] > 1.6], key=lambda t: t["y"])
    secs = []
    for k, sec in enumerate(ch["sections"]):
        hdr = None
        for b in sec["body"]:
            if b[0] == "table":
                hdr = b[1]
                break
        secs.append(parse_section_slide(slides[3 + k], hdr))
    app = _texts(slides[3 + len(ch["sections"])])
    app_en = next(t["text"] for t in app if abs(t["y"] - 1.20) < 0.03)
    return {"title_en": title_en, "step_note": note, "opening_en": oc_en,
            "opening_slide": [b["text"] for b in body[:2]], "opening_q": body[2]["text"] if len(body) > 2 else "",
            "sections": secs, "app_en": app_en}
