"""Convert a Ch01-Ch09 book docx (produced by the earlier generator) into a content dict.

The dict uses the same top-level fields as content/chNN.py; each section carries an
ordered `body` list so the book text round-trips exactly. Slide-only fields (English
titles, slide leads, slide block layout) are filled from the chapter's pptx by
convert_pptx.py.
"""
import re
import sys

import docx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def ptext(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def first_rpr(el):
    r = el.find(f".//{W}r/{W}rPr")
    return r


def classify(el, numfmt):
    tag = el.tag.replace(W, "")
    if tag == "tbl":
        rows = []
        for tr in el.findall(W + "tr"):
            rows.append([ptext(tc).strip() for tc in tr.findall(W + "tc")])
        widths = [int(g.get(W + "w")) for g in el.findall(f"{W}tblGrid/{W}gridCol")]
        return ("table", rows, widths)
    if tag != "p":
        return ("other", tag)
    txt = ptext(el).strip()
    ppr = el.find(W + "pPr")
    style = ppr.find(W + "pStyle").get(W + "val") if ppr is not None and ppr.find(W + "pStyle") is not None else ""
    if el.findall(f".//{W}drawing"):
        return ("image",)
    if el.findall(f".//{W}br") and not txt:
        return ("pagebreak",)
    if not txt:
        return ("blank",)
    if style == "Heading1":
        return ("h1", txt)
    if style == "Heading2":
        return ("h2", txt)
    if style == "ListParagraph":
        nid = el.find(f".//{W}numId")
        fmt = numfmt.get(nid.get(W + "val") if nid is not None else "", "bullet")
        return ("li", txt, fmt)
    shd = ppr.find(W + "shd") if ppr is not None else None
    if shd is not None and shd.get(W + "fill") == "FFF4DD":
        return ("callout", txt)
    r = first_rpr(el)
    color = r.find(W + "color").get(W + "val") if r is not None and r.find(W + "color") is not None else ""
    sz = r.find(W + "sz").get(W + "val") if r is not None and r.find(W + "sz") is not None else ""
    bold = r is not None and r.find(W + "b") is not None and r.find(W + "b").get(W + "val") not in ("false", "0")
    jc = ppr.find(W + "jc").get(W + "val") if ppr is not None and ppr.find(W + "jc") is not None else ""
    keep = ppr is not None and ppr.find(W + "keepNext") is not None
    if txt.startswith("مثال محلول:"):
        return ("ex", txt[len("مثال محلول:"):].strip())
    if color == "555555" and jc == "center":
        return ("caption", txt)
    if color == "555555":
        return ("small", txt)
    if color == "E09F3E" and sz == "26":
        return ("kicker", txt)
    if color == "0F4C5C" and sz == "52":
        return ("title", txt)
    if color == "0F4C5C" and bold and keep:
        return ("label", txt)
    if color == "0F4C5C" and bold:
        return ("teal", txt)
    if sz == "21":
        return ("ref", txt)
    if bold and sz == "26":
        return ("bold", txt)
    if sz == "23":
        return ("task", txt)
    return ("para", txt)


def numbering_formats(doc):
    out = {}
    try:
        num = doc.part.numbering_part.element
    except Exception:
        return out
    absfmt = {}
    for a in num.findall(W + "abstractNum"):
        lvl = a.find(f"{W}lvl[@{W}ilvl='0']")
        f = lvl.find(W + "numFmt").get(W + "val") if lvl is not None and lvl.find(W + "numFmt") is not None else "bullet"
        absfmt[a.get(W + "abstractNumId")] = f
    for n in num.findall(W + "num"):
        out[n.get(W + "numId")] = absfmt.get(n.find(W + "abstractNumId").get(W + "val"), "bullet")
    return out


def elements(path):
    d = docx.Document(path)
    nf = numbering_formats(d)
    return [classify(el, nf) for el in d.element.body if el.tag != W + "sectPr"]


AXIS_RE = re.compile(r"^المحور (\S+): (.+?) \((.+)\)$")


def convert(path, num):
    els = elements(path)
    i = 0
    ch = {"num": num}

    def take(kind):
        nonlocal i
        assert els[i][0] == kind, (path, i, kind, els[i])
        i += 1
        return els[i - 1]

    take("kicker")
    ch["title"] = take("title")[1]
    ch["question"] = take("callout")[1].replace("سؤال الفصل:", "", 1).strip()
    take("image")
    take("caption")
    h = take("h1")[1]
    assert h.startswith("حالة افتتاحية:"), h
    oc = {"title": h.split(":", 1)[1].strip(), "paras": []}
    while els[i][0] == "para":
        oc["paras"].append(take("para")[1])
    oc["question"] = take("callout")[1]
    ch["opening"] = oc
    assert take("h1")[1] == "مخرجات التعلم"
    take("para")
    outs = []
    while els[i][0] == "li":
        t = take("li")[1]
        m = re.match(r"^(\S+) \(([^)]+)\): (.+)$", t)
        outs.append((m.group(1), m.group(2), m.group(3)))
    ch["outcomes"] = outs
    assert take("h1")[1].startswith("خريطة الفصل")
    tbl = take("table")[1]
    axes = []
    for row in tbl[1:]:
        letter, name = row[0].split(". ", 1)
        axes.append({"letter": letter, "name": name, "q": row[1], "topics": row[2]})
    ch["axes"] = axes
    if els[i][0] == "blank":
        i += 1
    ch["output"] = take("callout")[1].replace("المخرج:", "", 1).strip()

    # sections
    sections = []
    axis = -1
    while True:
        kind = els[i][0]
        if kind == "h1":
            m = AXIS_RE.match(els[i][1])
            if not m:
                break
            axis += 1
            axes[axis]["en"] = m.group(3)
            i += 1
            continue
        h2 = take("h2")[1]
        mnum = re.match(r"^(\d+)\.(\d+) (.+)$", h2)
        sec = {"axis": axis, "t": mnum.group(3), "body": []}
        while els[i][0] not in ("h1", "h2"):
            e = els[i]
            i += 1
            if e[0] == "blank":
                continue
            if e[0] == "li":
                if sec["body"] and sec["body"][-1][0] == "list" and sec["body"][-1][2] == e[2]:
                    sec["body"][-1][1].append(e[1])
                else:
                    sec["body"].append(["list", [e[1]], e[2]])
                continue
            if e[0] == "table":
                sec["body"].append(["table", e[1][0], e[1][1:], e[2]])
                continue
            sec["body"].append(list(e))
        sections.append(sec)
    ch["sections"] = sections

    # application
    h = take("h1")[1]
    assert h.startswith("التطبيق:"), h
    app = {"title": h.split(":", 1)[1].strip(), "steps": []}
    while els[i][0] == "li":
        app["steps"].append(take("li")[1])
    meta = take("small")[1]
    m = re.match(r"الملف: (.+?)\s+\|\s+الزمن المقترح: (.+)$", meta)
    app["file"], app["time"] = m.group(1).strip(), m.group(2).strip()
    ch["app"] = app
    assert take("h1")[1] == "حالة للنقاش ومهمة بحثية"
    ch["discussion"] = []
    while els[i][0] == "li":
        ch["discussion"].append(take("li")[1])
    take("h2")
    ch["task"] = take("task")[1] if els[i][0] == "task" else take("para")[1]
    assert take("h1")[1] == "خلاصة الفصل"
    summ = []
    while els[i][0] == "teal":
        lab, txt = take("teal")[1].split(": ", 1)
        summ.append((lab, txt))
    ch["summary"] = summ
    ch["next"] = take("callout")[1]
    assert take("h1")[1] == "أسئلة المراجعة"
    rev = []
    while els[i][0] == "teal":
        t = take("teal")[1]
        m = re.match(r"^\d+\. \[([^\]]+)\] (.+)$", t)
        rev.append((m.group(1), m.group(2)))
    ch["review"] = rev
    assert take("h1")[1] == "المصطلحات"
    ch["terms"] = [tuple(r) for r in take("table")[1][1:]]
    if els[i][0] == "blank":
        i += 1
    assert take("h1")[1] == "المراجع"
    refs = []
    while els[i][0] in ("ref", "small", "bold", "para"):
        t = els[i][1]
        i += 1
        if " — " in t:
            a, b = t.split(" — ", 1)
        else:
            a, b = t, ""
        refs.append((a.strip(), b.strip()))
    ch["refs"] = refs
    take("pagebreak")
    assert take("h1")[1].startswith("ملحق")
    ans = []
    while i < len(els) and els[i][0] in ("bold", "para", "teal"):
        ans.append(re.sub(r"^\d+\.\s*", "", els[i][1]))
        i += 1
    ch["answers"] = ans
    rest = [e for e in els[i:] if e[0] != "blank"]
    assert not rest, rest[:3]
    return ch


if __name__ == "__main__":
    import json
    c = convert(sys.argv[1], int(sys.argv[2]))
    print(json.dumps({k: (v if k != "sections" else len(v)) for k, v in c.items()}, ensure_ascii=False)[:1500])
